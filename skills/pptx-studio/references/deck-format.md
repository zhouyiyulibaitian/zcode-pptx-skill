# deck.json 格式参考

`deck.json` 是整个工具链的唯一事实源：浏览器预览、浏览器编辑、PPTX 导出、PDF 导出都从它出发。
所有几何尺寸单位是**英寸**，字号单位是**点（pt）**，角度是**度**。

## 坐标系

- 画布固定为 `deck.size`：16:9 = `13.333 × 7.5`，4:3 = `10 × 7.5`。
- HTML 按 **96 dpi** 映射：`px = inch × 96`，字号 `px = pt × 4/3`。这与 PowerPoint 的比例一致，
  所以"浏览器里看到的换行/位置"和"PPTX 里的"是对应的，不是靠测量 DOM 猜出来的。
- 元素 `x/y/w/h` 是相对画布左上角的位置；`z` 决定叠放次序；`rotation` 顺时针为正。

## 顶层结构

```jsonc
{
  "schemaVersion": 1,
  "deck": {
    "title": "江南古典园林空间叙事研究",
    "size": {"w": 13.333, "h": 7.5, "preset": "16:9"},
    "theme": "ink",                    // 内置主题 id，或一个主题 JSON 的路径
    "themeOverrides": {                // 可选，局部覆盖主题令牌
      "colors": {"primary": "#1F4E79"},
      "fonts": {"body": {"ea": "等线"}},
      "geometry": {"bodySize": 15}
    },
    "language": "zh",
    "meta": {"author": "张三"}
  },
  "slides": [ /* 见下 */ ]
}
```

## 页面（slide）

一页有两种写法，**要么给版式，要么给显式元素**：

```jsonc
// 1) 版式驱动（推荐，改内容不改排版）
{"id": "s1", "layout": "cover", "content": {"title": "…", "subtitle": "…"},
 "notes": "讲稿", "footer": "答辩 · 2026", "chrome": true}

// 2) 显式元素（编辑器一旦拖动/改样式，这一页会自动转成这种形式并带 detached: true）
{"id": "s2", "elements": [ … ], "notes": "…"}
```

| 字段 | 说明 |
|---|---|
| `layout` | 版式名，见文末清单。有 `content` 时在渲染/导出时展开成元素 |
| `content` | 版式字段，见文末清单 |
| `elements` | 显式元素数组；存在时优先于 `layout` |
| `notes` | 演讲者备注，导出时写入 PPTX 的备注页 |
| `background` | `{"type":"solid","color":"#0E1B2E"}` / `{"type":"gradient","from":"#0E1B2E","to":"#1F4E79","angle":45}` / `{"type":"image","src":"assets/bg.png"}` |
| `transition` | `fade` / `push` / `wipe` / `cut`，写入 PPTX 切换效果 |
| `chrome` | `false` 可关掉页脚与页码 |
| `footer` | 覆盖本页页脚文字 |

页码与页脚在**渲染时生成**，不写回 deck.json，所以反复解析不会重复堆积。

## 元素通用字段

```jsonc
{"id": "s3-e1", "type": "text", "x": 0.9, "y": 1.0, "w": 6.0, "h": 0.8,
 "z": 10, "rotation": 0, "opacity": 1.0, "locked": false, "role": "title"}
```

`role` 只影响语义（`title` / `subtitle` / chrome 等），不改变渲染。`locked: true` 表示编辑器不可拖动（背景装饰用）。
`allow_bleed: true` 允许超出版心（满版图/背景）。

### text

```jsonc
{"type": "text", "x": 0.9, "y": 1.0, "w": 6, "h": 1.2,
 "paragraphs": [
   {"runs": [{"text": "加粗部分", "bold": true, "color": "#E23A2E"},
             {"text": "普通文字"}],
    "size": 16, "color": "#111111", "font": {"latin": "Arial", "ea": "微软雅黑"},
    "align": "left", "lineSpacing": 1.5, "spaceBefore": 0, "spaceAfter": 6,
    "bullet": false, "indent": 0}
 ],
 "valign": "top", "wrap": true,
 "fill": {"type": "solid", "color": "#F4F4F2", "alpha": 1.0},
 "line": {"color": "#D8D8D4", "width": 1, "dash": "dash"}}
```

- 段落级字段（`size/color/font/align/lineSpacing/bullet`）是默认值，run 级可覆盖。
- `align`: `left|center|right|justify`；`valign`: `top|middle|bottom`。
- `padding`: `[左, 上, 右, 下]`（英寸），默认 `[0.04,0.02,0.04,0.02]`。
- 一个段落想分多段 → `paragraphs` 里放多项；想同段不同样式 → 一个段落多个 `runs`。

### shape

```jsonc
{"type": "shape", "shape": "roundRect", "x": 0.9, "y": 2, "w": 3.4, "h": 2,
 "fill": {"type": "solid", "color": "#FFFFFF"},
 "line": {"color": "#DCE2EA", "width": 1},
 "radius": 0.12, "paragraphs": [ …可选，形状内文字… ]}
```

`shape` 取值：`rect` `roundRect` `ellipse` `line` `arrow` `triangle` `diamond` `chevron`
`pentagon` `star5` `parallelogram` `plus` `arc` `ring` `cloud`。
`fill.type` 可为 `solid` / `gradient`（`from`/`to`/`angle`）；`fill.alpha` 0–1 表示透明度。
`line.arrow: true` 给线加箭头。

### image

```jsonc
{"type": "image", "src": "assets/site.png", "x": 0.9, "y": 1.2, "w": 5.6, "h": 5.2,
 "fit": "cover", "radius": 0.1, "opacity": 1, "alt": "现场照片"}
```

- `src` 相对 deck.json 所在目录，或绝对路径。**不要用 `file://` 或网络地址**（导出时读不到）。
- `fit`: `cover`（裁切填满，PPTX 里用 crop 实现）/ `contain`（留白居中）。
- `src` 为空时导出为"图片占位"框，不会让构建失败。

### table

```jsonc
{"type": "table", "x": 0.9, "y": 1.6, "w": 11.5,
 "cols": [2.6, 4.0, 4.9],           // 列宽（英寸），省略则等分
 "rowHeight": 0.52,
 "header": true,                     // 首行按表头样式（主题色底、白字）
 "rows": [[{"text": "叙事要素"}, {"text": "空间对应"}, {"text": "设计手段"}],
          [{"text": "起"}, {"text": "入口引导段"}, {"text": "收窄视线"}]],
 "h": 2.6}
```

单元格可带 `fill` `color` `size` `bold` `align` `valign`。

### chart（导出为 PowerPoint 原生可编辑图表）

```jsonc
{"type": "chart", "chart": "column", "x": 0.9, "y": 1.6, "w": 7.2, "h": 4.4,
 "categories": ["2021", "2022", "2023"],
 "series": [{"name": "文献量", "values": [12, 19, 27]}],
 "legend": false, "dataLabels": true, "grid": true,
 "title": "图 1  近五年相关文献数量",
 "valueFormat": "0", "stacked": false, "yAxis": true}
```

- `chart`: `column` `bar` `line` `area` `pie` `doughnut` `radar` `scatter`。
- 每个系列 `values` 的长度必须与 `categories` 一致（校验会报错）。
- `valueFormat` 支持 Excel 风格（`0`、`0.0`、`#,##0`、`0%`）或 Python 风格（`{:.1f}`）。
- 导出到 PPTX 是**原生图表**，在 PowerPoint 里可以右键"编辑数据"继续改。

## 版式清单（layout → content 字段）

| 版式 | content 字段 | 适用 |
|---|---|---|
| `cover` | `title` `subtitle` `tag` `presenter` `org` `date` | 封面 |
| `cover-image` | `title` `subtitle` `image` `presenter` `date` | 半图封面 |
| `section` | `number` `title` `subtitle` | 章节过渡 |
| `toc` | `title` `items[{text, note}]` | 目录 |
| `bullets` | `title` `subtitle` `bullets[{text, sub[]}]` `footnote` `two_column` | 要点页 |
| `two-col` | `title` `left{title,bullets}` `right{title,bullets}` | 两栏对照 |
| `cards` | `title` `columns` `cards[{tag,title,body}]` | 卡片组（3–4 张） |
| `kpi` | `title` `metrics[{value,unit,label,delta}]` `footnote` | 关键数据 |
| `image-text` | `title` `image` `side`(left/right) `image_ratio` `body` `bullets` `caption` | 图文页 |
| `image-full` | `title` `image` `subtitle` `tag` `overlay` | 满版大图 |
| `quote` | `quote` `author` `source` | 引言 |
| `chart` | `title` `chart{…}` `insight` `bullets` `caption` | 图表页 |
| `table` | `title` `table{cols,rows,header}` `note` | 表格页 |
| `timeline` | `title` `steps[{label,title,text}]` | 时间线/流程 |
| `compare` | `title` `left{title,items}` `right{title,items}` | 前后对比 |
| `big-number` | `title` `value` `unit` `caption` `bullets` | 单一大数字 |
| `closing` | `title` `subtitle` `contact` | 结束页 |

版式会在需要时自动收缩字号以避免溢出；元素一旦被手动编辑，该页转为显式元素，不再跟随版式。

## 内置主题

| id | 名称 | 特征 |
|---|---|---|
| `swiss` | 瑞士网格 | 白底黑字、正红强调、无圆角、粗标题 |
| `ink` | 深墨商务 | 深墨蓝 + 暖金、圆角卡片、稳重（答辩/汇报默认） |
| `tech` | 科技渐变 | 深色 + 紫蓝渐变 + 青绿高亮 |
| `paper` | 纸感学术 | 米白纸底、衬线标题、墨绿/砖红 |

也可以把任意 PPTX 提取成主题文件后当作主题使用：
`python studio.py template 模板.pptx --out tpl.json` → `--theme tpl.json`。

主题 JSON 的结构就是 `assets/themes.json` 里的一项：`fonts`（title/body/mono 的 latin 与 ea）、
`colors`（bg/surface/text/muted/primary/accent/onPrimary/line）、`geometry`（margin/gutter/radius/
各类字号/行距）、`style`（标题线、卡片风格、页码开关等）。
