# 处理已有的 PPTX

## 三条路线，先选对

| 用户想要 | 路线 | 命令 |
|---|---|---|
| 只想看/了解这份 PPT 里有什么 | 解析 | `inspect`（`--outline` 出可读大纲，`--json` 出结构化数据） |
| 改文字、备注、页序，**保持原排版** | 原地编辑 | `edit`（保留原格式，只动指定内容） |
| 大改内容/重排/换模板，并想用浏览器预览编辑 | 转成 deck | `to-deck` → 浏览器编辑 → `build` |
| 换一套模板/统一视觉 | 套模板 | `apply-template` |

判断标准：**改动小于 30% 且要求"别动我的版式"→ 用 `edit`**；
要重排、要预览、要批量生成 → 用 `to-deck`/`apply-template`。

## inspect

```bash
python <skill>/scripts/studio.py inspect 汇报.pptx --outline          # 人读大纲
python <skill>/scripts/studio.py inspect 汇报.pptx --json a.json      # 机器读结构
```

`--json` 输出包含每页版式名、每个形状的类型/位置/文字/字号/颜色/字体、图片元数据（sha1、体积）、
表格内容、图表数据、备注，以及全局的 `fonts_used` / `colors_used` 统计。
**做视觉统一前先看 `fonts_used` 和 `colors_used`** —— 这是判断原稿用了什么字体和主色的最快方式。

## edit：在保留格式的前提下修改

```bash
python scripts/studio.py edit 原.pptx --out 新.pptx \
  --replace "旧文字=新文字" --replace "另一个=替换" \
  --notes-index 2 --notes "第二页讲稿" \
  --duplicate-slide 1 --delete-slide 6 --move-slide 4:1 \
  --extract-images images/
```

也支持操作文件：`--ops ops.json`，内容是一个数组：

```jsonc
[
  {"op": "replace", "find": "2024年", "replace": "2025年", "scope": "all"},
  {"op": "set_notes", "scope": "1,3-5", "text": "统一备注"},
  {"op": "add_slide", "layout": "标题和内容", "index": 3, "title": "新章节", "body": "要点一\n要点二"},
  {"op": "add_image", "slide": 2, "path": "a.png", "position": "cover"},
  {"op": "add_chart", "slide": 6, "spec": {"chart": "column", "categories": ["A","B"],
     "series": [{"name": "数量", "values": [3, 5]}], "y": 1.6, "h": 4.2}},
  {"op": "add_table", "slide": 7, "spec": {"rows": [[{"text":"表头"}],[{"text":"数据"}]]}},
  {"op": "set_font", "scope": "all", "font": "微软雅黑"},
  {"op": "set_background", "scope": "1", "color": "#0E1B2E"}
]
```

替换的边界行为：如果查找词**跨越多个格式片段**（PowerPoint 常把一个句子拆成多段 run），
替换会把结果合并到第一个 run 的格式上，并在报告里提示"已合并为单一格式"。这是有损的，
但比替换失败更有用；如果用户在意，改用 `to-deck` 路线。

## to-deck：把 PPTX 变成可预览可编辑的 deck

```bash
python scripts/studio.py to-deck 汇报.pptx --out 汇报.deck --theme ink
```

产出目录：

```
汇报.deck/
├── deck.json           # 唯一事实源
├── assets/             # 抽出的图片（按内容去重）
├── index.html          # 静态预览（可直接分享/打印）
└── index.html.edit.html # 可编辑版本，或直接 studio.py preview 起服务
```

转换时会：把 EMU 精确换算成英寸（几何不丢）、保留 run 级字号/颜色/加粗/字体（含 `a:ea` 中文字体）、
抽取图片、读取表格与图表数据、搬运备注、把满版背景图识别为背景层。

**转换不是 100% 还原**，这些会以 `warnings` 报出来而不是悄悄丢掉：

- SmartArt、3D 变形、艺术字、OLE 对象、媒体（视频/音频）无法解析 → 保留在原始文件里（`edit` 路线不动它们），
  但 `to-deck` 不会转换它们。
- 组合形状会**拍平**成绝对坐标的独立元素（位置正确，但失去组合关系）。
- 无法读取嵌入工作簿的图表 → 跳过并警告。
- 版式占位符继承的样式（未直接写在 run 上的字号/颜色）会退化成默认值。

## apply-template：把内容套到模板上

```bash
# 保留模板的母版/版式/配色/页眉页脚（学校模板、单位模板就用这个）
python scripts/studio.py apply-template 我的内容.pptx --template 学校模板.pptx \
       --out 答辩.pptx --workdir work --mode layout

# 只借模板的配色与字体，用内置版式重排（HTML 预览和 PPTX 完全一致）
python scripts/studio.py apply-template 我的内容.pptx --template 学校模板.pptx \
       --out 答辩.pptx --mode theme
```

两种模式的取舍：

- `--mode layout`（默认）：**PowerPoint 里最像模板**。幻灯片是模板自己的版式，母版上的校徽、
  装饰、页脚都在。但 HTML 预览按本工具的版式渲染，**预览不等于最终样式**（内容一致、装饰不同）。
- `--mode theme`：把模板的配色与字体提取成主题覆盖，用内置版式重排。**HTML 预览与 PPTX 一致**，
  适合"要预览、要反复改"的场景；代价是不带模板母版上的装饰图形。

要预览一致就用 `theme`，要"和模板一模一样"就用 `layout`。不要向用户含糊其辞，直接说明这个差别。

## 从模板提取主题

```bash
python scripts/studio.py template 学校模板.pptx --out 学校配色.json
python scripts/studio.py build --deck deck.json --out out.pptx --theme 学校配色.json
```

提取的是：出现频率最高的中文字体/西文字体、主色与强调色、正文字号中位数、母版背景色。
结果是一个普通主题 JSON，可以手工微调（比如把 `primary` 换成学校标准色）。

**把模板当 `--base` 更省事**：`build --base 学校模板.pptx` 直接沿用模板的母版、版式、
主题字体与配色，并有 `--layout-map` 把 deck 的版式映射到模板的版式名：

```json
{"cover": "标题幻灯片", "section": "节标题", "bullets": "标题和内容"}
```

模板自带的示例页会被自动清空，不会混进成品。

## 常见坑

| 现象 | 原因与处理 |
|---|---|
| 中文变成了宋体/等线 | 只设了 `a:latin`。本工具总是同时写 `a:ea`；自己写脚本时要注意 |
| 导出的 PPT 里图表是图片 | 用了 `image` 元素。要可编辑图表必须用 `chart` 元素 |
| 图片不显示/位置错 | `src` 用了绝对路径或网络地址；改成相对 deck.json 的 `assets/xxx.png` |
| PDF 每页多出空白页 | 版心尺寸与 `@page` 不一致。本工具用固定 `@page size`；自定义 CSS 时注意 |
| 字体在别人电脑上不一样 | 主题里用了本机才有字体。跨设备稳妥选择：微软雅黑 / 等线 / Arial / Georgia |
| `to-deck` 后排版变松 | 版式占位符继承的字号没读到，退化成默认值。可在 deck.json 里补 `size`，或在浏览器里调 |
