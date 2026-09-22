---
name: pptx-studio
description: 制作、修改、分析、导出演示文稿（PPT / 幻灯片 / deck）。当用户说「做个 PPT」「帮我做汇报/答辩/路演的幻灯片」「把这些内容整理成 PPT」，或上传、提到 .pptx 文件并想「看看里面有什么内容」「改一下文字/备注/页序」「分析一下它的排版配色字体」「套用这个模板」「换成我们学校的模板」「把 PPT 转成 HTML/网页预览」「导出成能继续编辑的 PPTX」「加个图表/备注」时，都使用本 skill。能力：读取解析已有 PPTX（文字/图片/表格/图表/备注/母版版式）、在保留原格式前提下批量修改、转成可在浏览器里预览与点击编辑的 HTML、按内置主题或现有模板生成新 PPT、导出原生可编辑 PPTX 与 PDF。
---

# pptx-studio

一套「HTML 可编辑工作流」的 PPT 工具：**内容写进 `deck.json`，浏览器里看效果并随手改，最后导出完全可编辑的 PPTX**。
也能反向工作：读已有的 PPTX、在保留原格式的前提下改、或者把内容搬到新模板上重排。

工具位置：本文件所在目录，命令入口 `<skill>/scripts/studio.py`（下称 `studio.py`）。
第一次使用前先跑一次自检：

```bash
python "<skill>/scripts/studio.py" themes    # 能列出主题就说明依赖正常
```

## 本机 DSH 适配规则（实测约束）

> 本节只记录**本机环境**约束。本 skill 的四处缺陷已在本机副本中修复，
> 所以下文不再有"绕开缺陷"的规则——按正文流程正常用即可。

**调用方式**

```bash
python = C:\Python314\python.exe   # 依赖装在系统 Python 的用户级 site-packages
命令：python "<skill>\scripts\studio.py" <子命令> ...   # <skill> = C:\Users\zhouyi\.agents\skills\pptx-studio
```

注意：**不要用 `python -s`，也不要设 `PYTHONNOUSERSITE=1`** —— python-pptx / Pillow /
pypdfium2 / lxml 都在用户级 site-packages 里，加这两个开关会让本 skill 直接报 ModuleNotFoundError。

**`pdf` 与 `shots` 在 DSH 会话里必然失败，用附带脚本代替**

原因：这两个子命令靠无头 Chrome/Edge 打印 HTML，而 DSH 的文件沙箱既拒绝 `Start-Process`
启动 Chrome（`Access is denied`），也拒绝 Chrome 多进程渲染所需的 mojo 命名管道
（`platform_channel.cc:108 拒绝访问 (0x5)`，退出码 -36863）。换 `--headless`、
`--single-process`、换 Edge、换 `--user-data-dir` 都无效——这是沙箱策略，**不是本 skill 的缺陷**
（同一份 HTML 在沙箱外导出正常）。

→ **改用本 skill 目录下的 `export-pdf.ps1`，在普通 PowerShell 窗口里执行**（不要经由 DSH 工具调用）：

```powershell
powershell -ExecutionPolicy Bypass -File "<skill>\export-pdf.ps1" -Deck "工作目录\deck.json" -Shots
```

- `-Deck` 会先调 `studio.py render` 生成 HTML 再打印成 PDF；也可用 `-Html` 指定已有 HTML。
- `-Out` 指定输出 PDF；`-Shots` 额外把每页渲染成 PNG（`-Scale` 默认 2.0）。
- 实测：14 页 deck → 351.7 KB PDF + 14 张 PNG，产物可被 pdf-toolkit 读回校验。

**质量门在不能出图时怎么过**（正文要求 `shots` 出图自己看，本机要换等价做法）

1. `validate` 必须 0 error —— 现在 `build` 自己会先跑 validate，**有 error 会直接中止**（exit 2，
   加 `--force` 才能硬导），所以这一条已经由工具保证，不会漏做；
2. 用 `export-pdf.ps1` 出 PDF 后，**用能读图的模型看 PNG**；若当前模型无图像输入，改用
   **页锚文本核对**：`python "<pdf-toolkit>\scripts\pdftool.py" text "<导出的.pdf>" --pages 1-20`
   逐页比对"页数 == deck 页数、每页标题正确、没有空页、没有异常截断"；
3. `validate` 报的"文字可能溢出"警告必须逐条处理。

**已修复的缺陷（本机副本已改，上游待 PR）**

| 原缺陷 | 修复后行为 |
|---|---|
| `edit --notes-index` 对无备注页静默空操作 | 现在按需创建备注页并写入；确实写不进去会报出是哪几页 |
| `inspect` 遇非 PPTX 抛 35 层 traceback | 进 python-pptx 前先验内容（`PK` 头 + 包结构），一行提示 + exit 2 |
| `to-deck` → `build` 重复叠加页码装饰 | 转换出的页带 `chrome: false`，重建不再生成第二份页码（实测形状数 164 → 164） |
| `validate` 漏检未知版式 / 非法 `schemaVersion` | 两者都报 error（未知版式会列出可用版式名） |
| `build` 不以 validate 为闸门 | 默认先 validate，有 error 即中止（exit 2）；`--force` 可绕过 |

**仍然存在的取舍（不是缺陷，交付时要向用户讲清）**

- `to-deck` 往返是**有损**的：图表标题会丢失、个别 run 被合并。转换后必须把 `warnings` 读出来告诉用户。
- `apply-template --mode layout` 与 `--mode theme` 的差别：前者在 PowerPoint 里最像模板但 HTML 预览
  展示本工具版式；后者预览与导出一致但不带模板母版装饰图形。不要含糊过去。

**实测基线（可拿来对照自己的结果）**

| 操作 | 实测结果 |
|---|---|
| `themes` / `layouts` / `demo` / `validate` / `build` / `render` / `inspect` / `resolve` / `template` / `to-deck` / `edit` | 全部 exit 0 |
| `build` 产出的 PPTX | 14 页、164 形状、**原生可编辑图表 1 个、原生表格 1 个**、3 页备注，无扁平化 |
| `preview` 服务 | `/api/health` 实测 **HTTP 200**；`POST /api/export/pptx` 返回 `{"ok":true}` |
| `export-pdf.ps1`（沙箱外） | 14 页 PDF 351.7 KB + 14 张 PNG，pdf-toolkit 读回为 `text-partial` |
| `pdf` / `shots`（沙箱内） | 必然失败 —— 用上面的脚本代替 |
## 先选路线

| 用户的话 | 路线 | 关键命令 |
|---|---|---|
| "做个 PPT""把这些内容做成幻灯片" | **生成** | `demo`/`init` → 写 `deck.json` → `validate` → `render`/`preview` → `build` |
| "这份 PPT 里有什么""分析一下" | **解析** | `inspect --outline` |
| "改一下文字/备注/删几页，别动排版" | **原地编辑** | `edit` |
| "重排/统一视觉/加图表，还要能预览" | **转模型再生成** | `to-deck` → 编辑 → `build` |
| "套用这个模板""用我们学校的模板" | **套模板** | `apply-template`（或 `build --base`） |
| "要能预览、还要能导出可编辑的 PPTX" | **生成** + `preview` | 见"浏览器编辑与交付" |

**判断要点**：改动小、且用户明确在意"别动我的版式" → `edit`；要重排、要预览、要批量 → 走 `deck.json`。
不要为了显得强大而把一份只需改两句话的 PPT 整份重建 —— 那会丢掉原排版。

## 质量门（每次交付前必须过）

1. `validate` 必须 **0 error**。有 `warn` 要说出来。
2. `render` 或 `preview` 后，用 `shots` 出图**自己看一眼**，确认没有溢出、遮挡、空白页。
3. 溢出警告（"文字可能溢出"）必须处理：缩短文案、换版式、或调整框大小。别把它当成噪音。
4. 交付时给出**实际产出文件的绝对路径**，不要只说"已经完成"。

```bash
python studio.py validate --deck deck.json
python studio.py shots   --deck deck.json          # 每页导出 PNG，自己看
```

## 工作流 A：生成新 PPT

### 1. 先定内容，再定视觉

先把内容敲定（每页一个论点），再写 `deck.json`。**一页只讲一件事**，标题写成结论句而不是名词短语。
页数默认 8–14 页；答辩/汇报建议：封面 → 目录 → 背景与问题 → 方法 → 结果/数据 → 讨论 → 结论 → 结束。

### 2. 写 deck.json

从示例开始最省事 —— `demo` 生成的 `deck.json` 覆盖了全部版式，直接改内容即可：

```bash
python studio.py demo --out work/deck.json --theme ink --title "选题名称"
python studio.py layouts      # 查看版式清单及其字段
```

最小可用结构（完整字段见 `references/deck-format.md`）：

```jsonc
{
  "schemaVersion": 1,
  "deck": {"title": "江南古典园林空间叙事研究", "theme": "ink",
           "size": {"w": 13.333, "h": 7.5, "preset": "16:9"}},
  "slides": [
    {"layout": "cover", "content": {"tag": "毕业设计", "title": "…", "subtitle": "…",
                                    "presenter": "答辩人：张三", "date": "2026 年 6 月"}},
    {"layout": "bullets", "notes": "讲稿写这里",
     "content": {"title": "研究背景", "bullets": [{"text": "…", "sub": ["…"]}]}},
    {"layout": "chart",
     "content": {"title": "研究现状",
                 "chart": {"chart": "column", "categories": ["2021","2022","2023"],
                           "series": [{"name": "文献量", "values": [12,19,27]}],
                           "dataLabels": true, "valueFormat": "0"},
                 "insight": "一句话结论", "caption": "数据来源：CNKI"}}
  ]
}
```

三条硬规则（违反会让导出结果变差，校验不一定报错）：

- **备注（`notes`）是讲稿不是重复页面文字。** 用完整句子写"这一页我要说什么"，这是答辩时真正救命的东西。
- **图表用 `chart` 元素，不要贴图片。** 只有 `chart` 才导出成 PowerPoint 原生可编辑图表。
- **数据要真实。** 不确定的数字要么标注来源与口径，要么留空让用户填；不要编造调研数据。

### 3. 校验、预览、导出

```bash
python studio.py validate --deck work/deck.json
python studio.py build  --deck work/deck.json --out work/答辩.pptx --transition fade
python studio.py pdf    --deck work/deck.json --out work/答辩.pdf
```

`build` 产出的 PPTX 里：文字是真文本（带 `a:latin`/`a:ea` 字体）、形状是原生自选图形、
表格是原生表格、图表是可"编辑数据"的原生图表、备注在备注页里 —— **没有任何东西被拍平成图片**。

## 工作流 B：改已有的 PPT

```bash
python studio.py inspect 原稿.pptx --outline                 # 先看清结构，别盲改
python studio.py edit 原稿.pptx --out 改后.pptx \
       --replace "2024年=2025年" --notes-index 2 --notes "…" \
       --delete-slide 6 --duplicate-slide 1 --extract-images imgs/
```

`edit` 只动用户点名的东西，其余格式原样保留；跨格式片段的替换会在报告里提示"已合并为单一格式"。
需要重排、换视觉、加图表、要预览时，走 `to-deck` 转成 `deck.json` 再生成：

```bash
python studio.py to-deck 原稿.pptx --out work/deck --theme ink
# → work/deck/deck.json + assets/ + index.html
```

转换是**有损**的（SmartArt/3D/艺术字/媒体不转，组合形状会被拍平，未直接写在 run 上的继承样式会退化）。
转换后**必须**把 `warnings` 读出来告诉用户，不要假装无损。

## 工作流 C：套模板

```bash
# 要"PowerPoint 里和模板一模一样"（保留模板母版/版式/校徽页脚）
python studio.py apply-template 内容.pptx --template 学校模板.pptx --out 答辩.pptx --mode layout

# 要"HTML 预览和 PPTX 完全一致"（借用模板配色字体，用内置版式重排）
python studio.py apply-template 内容.pptx --template 学校模板.pptx --out 答辩.pptx --mode theme

# 或者：已有 deck.json，直接以模板为基础构建
python studio.py template 学校模板.pptx --out tpl/theme.json     # 提取配色字体
python studio.py build --deck deck.json --out out.pptx --theme tpl/theme.json
python studio.py build --deck deck.json --out out.pptx --base 学校模板.pptx \
       --layout-map tpl/layout-map.json
```

`--mode layout` 与 `--mode theme` 的差别必须向用户讲清楚：前者 PowerPoint 里最像模板，但 HTML 预览展示的是
本工具的版式（内容一致、装饰不同）；后者预览与导出一致，但不带模板母版上的装饰图形。**不要含糊**。
模板自带的示例页会被自动清空，不会混进成品。

## 工作流 D：浏览器编辑与交付

```bash
python studio.py preview --deck work/deck.json --port 5390   # 后台起服务
```

- 编辑界面 `http://127.0.0.1:5390/`：点选元素拖动/改尺寸、双击改文字、右侧面板改字号颜色对齐、改备注、
  增删复制页面、切主题、换版式、替换图片、溢出检查、底部"导出 PPTX / 导出 PDF"。
- 放映预览 `http://127.0.0.1:5390/preview`：← → 翻页、`F` 全屏、`P` 打印。
- 用户改完会写回 `deck.json`，**静态 HTML 会同步刷新**；agent 继续改 `deck.json` 时刷新浏览器即可看到。

交付物按需给：`deck.json`（可继续改）、`deck.html`（静态预览，双击可看）、`out.pptx`（可编辑成品）、`out.pdf`。

## 常见坑

| 现象 | 处理 |
|---|---|
| 中文显示成宋体/等线 | 导出时 `a:ea` 已写入，若仍不对说明主题里指定了本机没有的字体 |
| 图片不显示 | `src` 必须是相对 `deck.json` 的路径（放 `assets/`），不要绝对路径/网络地址 |
| 文字溢出 | `validate` 会警告；版式本身会自动缩字号，手改过尺寸的页面需要手动收 |
| PDF 多出空白页 | 版心尺寸要和 `@page` 一致；本工具已处理，自定义 CSS 时注意 |
| 换台电脑字体变了 | 稳妥字体：微软雅黑 / 等线 / Arial / Georgia；避免用只有本机装的字体 |
| 图表在 PPT 里不能改数据 | 用了 `image` 而不是 `chart` 元素 |

## 深入参考

- `references/deck-format.md` —— deck.json 全字段、元素类型、17 种版式、主题令牌
- `references/existing-pptx.md` —— 解析/编辑/转换/套模板的细节、保真度边界、故障排查
- `references/workflow-qa.md` —— 自查清单与交付话术
