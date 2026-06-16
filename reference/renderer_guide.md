# Renderer HTML 编写指南

renderer.py 的使用说明书。写 HTML 时必须遵守这些规则，否则 renderer 无法正确转换。

---

## 画布

1920x1080px。所有内容在一个 div 内。overflow:hidden。

## 调色板 -- 严格：只用这 9 个色值

- #305496 深蓝：thin nav bar bg, slim takeaway strip bg
- #4472C4 主蓝：图表柱子, 关键指标数字, 左侧强调线
- #5B9BD5 次蓝：次级图表系列, subtle emphasis
- #D9EAF7 浅蓝：局部高亮底色, table header tint, phase marker
- #EAF1F8 面板蓝：少量分组面板 tint，仅在版式需要时使用
- #B4C7E7 线框蓝：边框, 分隔线, 网格线
- #FFFFFF 白：slide bg, 卡片 bg（主导色）
- #1C1917 暖黑：标题文字, 正文文字
- #666666 灰：坐标轴, source 行, 脚注

不允许其他颜色。不要发明近似色如 #f2f5f9。

## 色彩平衡

白色是主导色。深蓝 #305496 只用于窄元素：nav bar (~36-42px), header band (~25-30px), takeaway strip (~36-48px), 卡片左侧强调线 (4px)。
不要用深蓝填充大面积 (>100px 高)，也不要用浅蓝给整块内容区或多个横向 band 大面积铺底。卡片默认用 #FFFFFF + #B4C7E7 边框；#EAF1F8 只用于少量重点 callout、表格分组或模板明确要求的分析对象。图表在白色背景上。

## 字体

英文优先 Arial。中文优先 `Microsoft YaHei`，macOS 可用 `PingFang SC`，并保留 Arial fallback。
不要使用稀有字体、手写字体、图标字体或 emoji 字体。
颜色：标题 #1C1917，正文 #1C1917，指标数字 #4472C4，source #666666。

推荐 CSS：

```css
body {
  font-family: "Microsoft YaHei", "PingFang SC", Arial, sans-serif;
}
```

## 布局结构 -- flex 自适应，不手动估坐标

**不要用 position:absolute 手动写 left/top 坐标。** 手动估坐标一定会估错，导致内容太小或位置偏移。让浏览器通过 flex 自动计算位置，renderer 读取浏览器算出来的结果。

slide 根元素是 flex column，width:1920px，height:1080px，overflow:hidden。按顺序有 5 个子元素：

1. Nav bar — flex-shrink:0, 固定高度, 全宽深蓝；视觉上是目录条，不是大块 tab 按钮
2. Title area — flex-shrink:0, 固定高度, 白底 + 分隔线
3. CONTENT WRAP — **flex:1 1 auto; min-height:0** — 自动填充所有剩余高度
4. Takeaway strip — flex-shrink:0, 固定高度, 全宽深蓝；文字优先，不默认配 icon
5. Source line — flex-shrink:0, 固定高度, 小灰字

**关键：** content wrapper 用 flex:1 自动吃掉 nav/title/takeaway/source 之外的所有空间。内容在 wrapper 里面也用 flex 分配空间（flex:1 拉伸、gap 控制间距），不手动写像素高度。

**这样做的好处：** 不管 nav 是 50px 还是 70px，不管 title 是一行还是两行，content 区域永远自动填满剩余空间，不会出现底部大片空白。

## 卡片内部分布

当卡片被 flex:1 拉高时，文字不要堆在顶部留下底部空白。
每张卡片自身也是 flex column + justify-content:space-between。
超过 ~200px 高的卡片必须有 >=3 个纵向分布的内容块。

## 密度

content wrapper 区域必须填充 >=90%。如果元素少，让每个元素更高（flex:1 拉伸），不要留空白带。

## 首选视觉语言：Paopao 简练咨询风格

默认页面应接近用户给出的参考风格，但只统一视觉语言，不统一版式排布。版式结构必须跟随选中的 prompt annotation，不能把所有页面强行改成同一种 SCR、矩阵、侧栏或右栏结构。

- 顶部一个大号黑色粗体 action title，1-2 行，避免小标题堆叠。
- 白色主背景，深蓝用于强调条/选中态/关键标识，浅蓝只用于局部高亮或必要分组，灰色用于次级线条和说明文字。
- 线条和线框要细、准、少：避免厚重边框、随机多余外框、虚线引导框、装饰性分割线、重复嵌套卡片框。
- 卡片、表格、图表、callout 都应像简洁可编辑 PPT 素材：默认白底，#B4C7E7 / #4472C4 细边框，圆角不超过 8px；只有重点 callout、表头或必要分组使用 #EAF1F8 浅蓝底。
- 底部如果有 takeaway strip，用深蓝平面细条，左侧放 `KEY TAKEAWAY`/`核心结论` 文本标签，右侧放一句综合判断；不要默认加灯泡、目标、箭头等 icon。
- 图标是允许的，但不是默认装饰。优先用文字标签、编号 badge、细分隔线、状态点、Harvey ball 或 KPI 标记表达结构；只有在能增加明确语义、或参考图/模板明确要求时才使用图标。若使用图标，应少量、简单、线性、蓝/灰双色，能用基础形状或小局部资产复刻；不要给每个卡片/每条 bullet/每个 band 都配图标，不要复杂插画、3D、emoji 感图标或大面积图标装饰。
- 避免照片背景、复杂纹理、渐变、玻璃拟态、重阴影、发光、装饰性背景和过多色彩。
- 页面可密，但必须规整：所有模块边缘对齐，分隔线统一，信息层级清楚。
- 避免大型装饰图、渐变背景、照片、3D、玻璃拟态、大面积阴影、复杂纹理。

## 图表 -- 关键

用 `<div data-chart="column|line|doughnut|scatter" data-categories='[...]' data-series='[...]'>` 写可编辑图表。

**绝对禁止内联 SVG。** renderer 会拒绝所有 SVG 元素。需要 icon 时优先用可编辑基础形状重画；若确实需要保留 SVG 风格 icon，先转成小 PNG 资产并用 `data-pptx-image` 标注。
**绝对禁止用 CSS height 模拟柱状图。** 用 data-chart 属性。

data-chart 示例：

```html
<div data-chart="column"
     data-categories='["2020","2021","2022","2023","2024"]'
     data-series='[{"name":"Revenue","values":[100,120,150,180,210]},{"name":"Cost","values":[80,90,100,110,120]}]'
     style="width:600px;height:300px;">
</div>
```

支持的图表类型：column, line, doughnut, scatter

## 表格

用标准 `<table>`，表头 #D9EAF7 背景，#B4C7E7 边框。

## 保留图片资产（仅用于小的复杂区域）

不要把整页 Image2 做背景。不要把整页截图/PNG/JPG/PDF 铺成底图。不要把整页图片底图作为任何方案建议。小的保留资产（logo、复杂图标）允许，但必须标注。

对外说明时不要说"无图片"，因为 icon 可能就是局部 PNG。正确说法是："没有整页图片底图；文本和主要形状可编辑，局部复杂 icon 可作为小图片资产保留。"

```html
<div data-pptx-image="assets/R01_name.png"
     data-pptx-preserve="true"
     data-pptx-preserve-reason="具体原因">
</div>
```

不要用 CSS background-image。不要用没有标注的 `<img>`。

### Icon 复刻规则

Image2 参考图里的 icon 是版面的一部分，不是装饰噪音。写 HTML 时必须保留其视觉语义。

- 如果参考图中有天平、循环箭头、日历、脑、图表、人物、齿轮、灯泡、盾牌、机器人、地球、火箭、警告、购物车、路标、目标等 icon，HTML 中不能用字母或缩写代替。
- `ECO`、`SYS`、`BANK`、`PROD`、`NET`、`UP`、`CASE`、`USER`、`STD` 等短文字只能用作真正的文字标签，不能放进 `.icon/.ibox/.circle/.mini/.ricon` 之类容器假装 icon。插件检查会把这种写法判为失败。
- 默认做法：优先用可编辑基础形状/边框/线条/简单字符重画 icon，使 PPTX 中 icon 仍可编辑。
- 备选做法：把 icon 作为小 PNG 资产保存到当前任务的 `assets/` 目录，并用 `data-pptx-image` 标注保留原因。
- 从 Image2/截图裁切 icon 时，必须先做清理：裁切框向外留安全边、去掉四角识别出的底色、重新按真实图形 trim 并导出透明 PNG。不要把带白底、蓝底、浅蓝底、相邻文字残片、分隔线残片的截图直接放进 `html/assets/`。
- 推荐命令：

```bash
python3 <plugin-root>/scripts/paopao_run.py clean-icon-crop \
  --image output/<task>/image2/image2_reference_01.png \
  --box 120,240,64,64 \
  --output output/<task>/html/assets/slide01_icon_target.png
```

  `--box` 是源图像里的 `x,y,w,h`。命令会默认向外扩 14px 防止裁不完整，自动移除四角底色并输出透明 PNG。如图形仍被裁断，优先加大 `--expand` 或重新取更宽的 `--box`，不要交付残缺 icon。
- `data-pptx-image` 可以和普通自动抽取 HTML 混用；它只额外保留小图片资产，不应导致正文、表格、形状丢失。
- 允许替代做法：用可编辑基础形状重画 icon，但 PPTX 预览里必须能被识别为同一语义。
- 如果图片 icon 在 PPTX 内部存在，但实际预览只显示浅色圆底、白块或空框，不能算通过；改用可编辑符号/基础形状。
- 如果小 PNG 裁切 icon 带入了相邻文字、分隔线、深蓝底栏、白色/浅蓝遮罩块、或任何可见底色，不能算通过；用 `clean-icon-crop` 重新裁切清理，或改成可编辑图形/统一专业 icon 资产。
- 如果参考图本身是编号 badge 或字母 badge，才可以保留为文本。
- QA 时若参考图有 icon 而 PPTX 内部 `pictures=0`，且 HTML 也没有等价可编辑图形，判定为失败。即使 `pictures>0`，也必须看最终预览确认 icon 细节可见。

示例：

```html
<div class="iconbox"
     data-pptx-image="assets/slide01_rollover_icon.png"
     data-pptx-preserve="true"
     data-pptx-preserve-reason="Icon from Image2 reference; semantic symbol for rollover risk.">
</div>
```

## Image2 复刻模式

当 HTML 是从 Image2 参考图复刻时：
- 复刻参考图的布局，不要在同一主题上重新设计
- 保留参考图的模块数量、相对几何、nav/title/footer 位置、图表/卡片/表格排列、密度和留白
- HTML 的依据是重新打开后真实看到的 Image2 参考图，而不是 final prompt、analysis_report、主题记忆或上一次脑补的布局。Prompt 决定生成参考图，参考图决定 HTML/PPTX。
- 参考图中每个主要可见模块在 HTML 中都必须有对应的可编辑元素
- 如果后续 PPTX 截图看起来不像参考图，即使可编辑也算失败
- 如果 HTML 第一版已经不像参考图，不要继续渲染 PPTX 期待后面修补；先 refill HTML/assets。
- HTML 前必须有 `output/spec/slideXX_spec.md`。spec 负责把 Image2 的可见元素、grid、组件、icon plan 转成结构化任务。HTML 必须按 spec 落地，不能直接凭感觉写。
- spec/HTML/PPTX 都不能包含整页底图；只有小范围 icon/logo/复杂局部资产允许保留。
- 渲染后必须重新打开对应 Image2 与最终 PPTX/实际预览逐页对照。不要凭 prompt、记忆、HTML视觉效果或“我觉得差不多”交付。参考图是视觉合同，PPTX 是交付物，两者必须一页一页核对。
- 每页对照至少检查：标题行数/重量、nav位置和active状态、主模块数量和几何比例、卡片密度、图标存在和语义、颜色层级、边框粗细、底部takeaway、source行。如果PPTX只是“大概同风格”而不是“结构跟参考图对得上”，判定失败。
- 对照完成后写 `qa/fidelity_review.json`，不要写 Markdown 版对照报告。该JSON只是内部交付闸门，不向用户展示。
- `qa/fidelity_review.json` 每页的 `dimensions_checked` 必须显式包含 `nav`、`title`、`module_geometry`、`icons`、`takeaway`、`color_hierarchy`。其中 `icons` 可记录为“无不必要图标/必要图标已保留”。缺任何一项都不能交付。

### Refill 源层闭环

当发现 HTML/PPTX 不像 Image2 时，标准动作是回到源层重新填充，而不是让最终 PPTX 成为第二套设计源：

- Image2 本身不适合转 PPT：重新生成 Image2，要求更简单、紧凑，并减少不必要 icon。
- HTML 不像 Image2：重写 HTML 和小图资产，重新跑 HTML check。
- PPTX 不像 HTML/Image2：改 HTML 稳定写法或 renderer，再重新渲染。
- 交付目录存在多个草稿 PPTX：只保留标准交付文件，把草稿移到内部目录，避免用户打开旧版。

### 回头看图规则

生成 Image2 后，不允许凭 prompt 记忆或数据摘要写 HTML。每页 HTML 开始前，必须重新打开当前页参考图，并按 nav / title / content / charts-tables-complex-geometry / takeaway / source 六个区域逐区观察。

写 HTML 前必须确认：
- nav 的 tab 数量、active 位置、分割线、三角指示是否看清
- title 的原文、行数、下划线、上下间距是否看清
- content 的一级模块数量、相对位置、宽高比例、内部元素是否看清
- 图表、表格、漏斗、箭头链、矩阵、流程图等复杂结构的数量和方向是否看清
- takeaway 的高度、左侧 label、竖向分割线、右侧正文行数是否看清
- source 的位置、字号、颜色是否看清

写完 HTML 后，必须再次打开同一张参考图，对照上面的六个区域修正明显偏差。这个步骤是为了让第一版 HTML 尽量按真实图复刻，不是为了反复无限循环。

### Image2 可用性门槛

下列 Image2 参考图必须废弃并重新生成：

- 中文或英文主文字明显乱码、缺字、混乱、不可辨认。
- 页面过度稀疏，像海报或封面，不像可编辑咨询页。
- 不是 16:9 横版画布，尤其是 3:2、4:3、竖版或方图。
- 视觉语言偏离参考：颜色不在蓝/白/灰/黑体系，背景有复杂纹理/渐变/照片，线框厚重或多余，卡片像模板堆砌而不是简练可编辑 PPT 素材。
- 使用照片、复杂产品图、人物图或全页背景作为主要视觉，但用户没有要求保留这些视觉。
- 图标复杂到无法用基础形状或小局部资产复刻。
- 颜色超出蓝/白/灰/黑主体系，或出现明显红/绿/黄/橙/紫/青大面积强调。
- 模块边界模糊、卡片不齐、文字贴边、层级混乱，导致无法稳定转换为 PPT。

### 高损失区域写法

以下区域最容易造成 Image2 → HTML 差异，必须按固定组件思路写：

- Icon：参考图中语义必要的 icon 必须用小 PNG 资产或可识别的可编辑图形复刻，不能替换成字母/缩写；若参考图出现大量装饰性 icon，应优先回到 Image2 重新生成更自然的低 icon 版本。
- Takeaway：固定高度、左侧 label、竖向分割线、右侧 1-2 行正文。不要让正文换行失控，不要省略分割线。
- 漏斗：用层级 band 组件，层数、每层宽度递减、编号位置、右侧注释标签必须与参考图一致。
- 箭头链 / chevron：用段落组件，段数、方向、active 高亮段、每段内部标题/KPI/说明必须与参考图一致。
- 矩阵：用标准 `<table>`，行列数、表头、侧边栏、圆点评级、右侧优先级栏不能少。
- KPI 卡：数字、单位、主说明、次说明、左侧强调线/编号/边框必须完整复刻。

### PPTX 转换稳定写法

浏览器能靠 flex 居中，但 PPTX 里的文字框不一定继承 flex 的横向和纵向居中。写 HTML 时必须把这些视觉意图显式写出来：

- PPTX QA：`renderer.py --pdf` 是 HTML/browser 预览，不等于 PowerPoint 实际打开效果。最终判断必须直接看 PowerPoint 实际打开的 PPTX。不要把 LibreOffice/PDF/PNG 转换作为默认 QA 路径，也不要要求用户为 QA 安装 LibreOffice。实际 PowerPoint 复核完成后必须写 `qa/powerpoint_review.json`，并通过 `paopao_run.py check --stage pipeline`。
- Nav：视觉上必须是薄目录条，不是多个大色块 tab。父级 nav 使用 `background:#305496`；子项使用透明背景、白字、紧凑间距、点状/细线分隔；active 项用加粗、下划线或小面积 #4472C4 标记。不要把每个 nav item 都画成等宽蓝色按钮。
- Nav 命名：导航容器必须使用 `<nav>` 或 `.nav/.navbar/.navigation/.tabs/.tabbar/.breadcrumb`，tab 子元素必须使用 `.tab` 或同类语义 class。renderer 会把这些元素内的文字识别为 `navText`，在 PPTX 写出时强制垂直居中。不要把导航栏写成没有语义 class 的普通 `div`。
- Nav QA：不能只看 HTML 截图。必须检查 PowerPoint 编辑界面或 PPTX 内部形状，确认每个 tab 的文字水平/垂直居中，非 active tab 不是白色填充。
- Takeaway：左侧 label 和右侧正文容器都必须显式写背景色；深蓝 takeaway 内部不能出现没有背景的子容器，否则 PPT 可能生成白色覆盖块。高度应保持细条感，通常 36-48px；不要为了“系统感”做成高横幅。
- Badge + title：不要用 inline badge 紧贴 inline 标题。用两列 grid 或 flex row：左列固定 badge 宽度，右列标题独立容器，避免 PPT 中编号和标题重叠。
- Short labels：导航、编号、KPI 数字、表头、按钮式标签必须显式写 `text-align:center`，并尽量使用固定高度容器。
- Tables and takeaway：表格区域和 takeaway 之间必须保留明确 gap；表格不要用 `flex:1` 撑到页面底部，优先给表格区域固定或最大高度。
- 空白控制：content 与 takeaway 之间的空白不能超过常规 gap，除非参考图明确存在。若出现大空白，优先增加表格高度、增大主体模块、或上移 takeaway。
- 框线控制：不要为了排版生成无意义左右空框。每个边框都必须对应参考图中的可见模块。正文文字不需要框时，不要额外包一层带边框容器。
- QA 分类：如果 HTML 正常但 PowerPoint 编辑界面错位、白块覆盖、文字靠左上，这是转换稳定问题；要改 HTML 稳定写法或 renderer，不要只重复生成。
- 数字符号：关键数值和目标值不要依赖 HTML entity 或复杂 inline 结构。`<3.4` 这类阈值在正文中优先写成全角 `＜3.4`，或把比较符、数字、单位拆成独立文本块；不要把 raw `<` 放进 HTML 文本节点。
- 指标块：大号数字、单位、说明文字不要嵌套 `<small>` 或多层 inline 元素。用独立块级元素，例如 `metric-label`、`metric-num`、`metric-unit`，并给每块固定高度/行高，避免 PowerPoint 丢失或裁掉数字。
- 形状稳定：关键箭头、chevron、漏斗、警示框不要依赖 `clip-path`、复杂 transform 或 CSS-only 三角组合。优先用矩形、边框、简单伪元素，或直接换成稳定的卡片/分段结构。
- 边框稳定：大面积虚线、细碎装饰线在 PowerPoint 里容易显得脏或漂移。除非参考图明确要求，优先用实线、分隔线和浅蓝底卡片。
- 变换风险：`transform: rotate(...)`、复杂 CSS transform、clip-path、emoji/symbol 字体、native table 样式，在 PowerPoint 中可能和 HTML 预览不同。需要用 PPTX 实际画面验证，不能只看 browser/PDF。

## 质量底线

以下是硬失败：
- 黑色或接近黑色的图表柱子。图表系列只能用 #305496, #5B9BD5, #4472C4, #666666
- 主标题、正文、标签、source、takeaway 乱码、缺字、不可读，或中文字体替换异常
- 卡片文字堆在底部。内容从顶部开始，有层级
- 三列纯文字加大面积空白。加指标条、证据行、mini-bar
- 卡片内大面积空白。高卡片需要 top/middle/bottom 三层内容
- 子弹堆砌。长段落转成分组要点 + 粗体引导句 + 底部行动行
- PowerPoint 编辑界面中导航文字靠左上、蓝底被白块覆盖、编号压住标题、表格压住 takeaway、content 与 takeaway 中间出现无意义大空白
- 整页图片底图、未标注 `<img>`、CSS `background-image`、内联 SVG、或任何导致主要文字不可编辑的做法
- 参考图有 icon，但 PPTX 中 icon 不可见、变成空框，或语义不再可识别
