# dsh-docx-from-images

把截图 / 扫描件里的试卷、讲义重新排成 **排版一致的 Word 文档** —— 这是 DeepSeek Harness 的一个 Host 插件（bundle），
提供两个工具和一份技能（skill），并把 OOXML 排版里最容易踩的几个坑全部内置规避。

> Screenshot → Word typesetting plugin for DeepSeek Harness. Two tools (`docx_typeset`, `docx_crop`) and one skill (`image-to-docx`).

![示例](docs/example.png)

## 提供的能力

| 名称 | 类型 | 作用 |
| --- | --- | --- |
| `docx_typeset` | Tool | 按 JSON spec 生成 `.docx`：着重号、制表位列（2/4 栏选项）、色带标题、**表格**、**插图**、显式分页；生成前会估算列宽并返回 warning |
| `docx_crop` | Tool | 从截图里裁插图：像素框裁剪、去白边、放大，或先输出带坐标的网格预览便于选框 |
| `image-to-docx` | Skill | 端到端流程：探图 → 逐图转写 → 标格式 → 写 spec → 生成 → 目视验收 → 交付话术 |

## 安装

需要本机有 Python 3 + `python-docx` + `Pillow`。插件会依次尝试：
`$DSH_PYTHON` → `$DSH_HOME/dsh-runtimes/*/dependencies/python` → `$HOME/.dsh/...` → PATH 上的 `python`。

```powershell
# 把仓库放到任意目录，然后在 Harness 里用 plugin_manager 安装这个目录
plugin_manager({ action: "install_bundle", target: "<本仓库绝对路径>" })
```

装好后：

- 工具出现在 Agent 的工具列表里（`docx_typeset` / `docx_crop`）；
- 技能目录里多出 `image-to-docx`，用户说“把这几张截图排成 Word”时会自动加载；
- 插件面板里的标题与图标来自 `locale/zh.json`、`locale/en.json` 与 `icon.svg`。

## 用法

```
docx_typeset({
  spec: {
    out: "D:/out/试卷.docx",
    page: { margin_cm: { top: 2.2, bottom: 2.2, left: 2.2, right: 2.2 } },
    body: { font: "宋体", latin: "Times New Roman", size_pt: 12, line_pt: 20 },
    blocks: [
      { type: "text", text: "1.\t下列加点字读音全部正确的一项是{b}（　）{/b}", tabs_cm: [0.78], before_pt: 7 },
      { type: "text", text: "A.\t屹{d}立{/d}（yì）\t霎{d}时{/d}（shà）", tabs_cm: [6.05, 9.4], indent_cm: 1.1 }
    ]
  }
})
```

块类型：`title` / `text` / `table` / `image` / `pagebreak`。
文本标记：`{d}着重号{/d}`、`{b}加粗{/b}`、`{i}斜体{/i}`、`{u}下划线{/u}`；`\t` 开一个新列。
完整示例见 [`examples/spec.json`](examples/spec.json)。

### 混排与页脚

一份卷子里常常“题干宋体、材料楷体、标题微软雅黑”，所以每个 `text` / `title` 块都能单独指定字体：

```json
{ "type": "text", "text": "①夜深了，人们睡了……", "font": "楷体", "first_line_cm": 1.06 }
```

页脚用 `{page}` / `{pages}` 占位，写入的是真正的 PAGE / NUMPAGES 域（合并多张图后页数自动正确）：

```json
"footer": { "text": "第 {page} 页 共 {pages} 页", "size_pt": 9, "align": "center" }
```

`body.dot_min_line_pt`（默认 24）控制带着重号的行至少多高；字号偏小时（≤13pt）调到 21–22 既能保住点，又不会把版面撑到多出一页。

## 内置规避的坑

| 坑 | 现象 | 本插件的做法 |
| --- | --- | --- |
| 着重号画在字上 | `w:em val="dot"` 把点画在字顶 | 强制 `val="underDot"`；也不拿“下划线”冒充 |
| 着重号被裁一半 | 行距“固定值”时字下方的点被切掉 | 含 `{d}` 的行自动升到 `body.dot_min_line_pt`（默认 24pt） |
| 字号拍脑袋 | 定大一号就每题多折一行，9 页排成 13 页 | 从截图反推单字宽（行宽 ÷ 字数），再以“分页是否与原图一致”验收 |
| 材料与题干同字体 | 楷体引文排成宋体一眼假 | 段落级 `font` / `latin` 覆盖 `body` |
| 页脚写死页数 | “共 27 页”在合并稿里全错 | `footer` 里的 `{page}` / `{pages}` 写成域代码 |
| 页边一排 ■ | 段落设了“与下段同页/段前分页/段中不分页” | 一律不写 `keepNext/keepLines/pageBreakBefore`；改在上一段末尾插分页符 |
| 表格后排不掉页 | 分页符落在表格前的段落上 | 表格后分页用 1pt 载体段承载 |
| 选项整行右移 | 行内标签 `A.\t` 只替换了第一个制表符 | 开头的 `A.\t` 自动归一成 `A. ` |
| 制表位叠加缩进 | 以为 `tabs_cm` 从首行缩进算起 | 制表位从**页面左边距**量起，并估算列宽给出溢出 warning |
| 插图糊 / 占满版心 | 直接塞整页截图 | 先用 `docx_crop` 裁图（去白边、2x 放大），再按版心比例给 `width_cm` |
| 表格手工空格对齐 | 列不齐、跨页表头丢失 | `Table Grid` + 显式列宽 + 表头跨页重复 |

## 验证（可选，需要装有 Word）

```powershell
$w = New-Object -ComObject Word.Application; $w.Visible = $false
$d = $w.Documents.Open($docx, $false, $true)
$d.ExportAsFixedFormat($pdf, 17); $d.Close(0); $w.Quit()
```

再用 `Windows.Data.Pdf` 逐页 `RenderToStreamAsync`（`DestinationWidth = 1240`）出 PNG 目视比对。
注意 PowerShell 5.1 会把 `.ps1` 当 ANSI 读，脚本内容与路径保持纯 ASCII，中文路径用参数传。

## 目录结构

```
index.js        插件入口：注册 docx_typeset / docx_crop / image-to-docx
builder.py      JSON spec -> .docx（python-docx）
crop.py         截取插图 / 网格预览 / 尺寸探测（Pillow）
cordis.patch.yml  bundle 安装行
locale/{zh,en}.json  插件面板的标题与描述
icon.svg        插件图标
examples/       可直接运行的示例 spec 与插图
```

## License

MIT © wjx-ai