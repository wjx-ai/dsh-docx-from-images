/**
 * Screenshot -> Word typesetting.
 *
 * Provides two mechanical tools and one skill:
 *   docx_typeset - build a .docx from a small JSON layout spec (着重号 / tab columns /
 *                  tables / pictures / explicit page breaks), via python-docx
 *   docx_crop    - crop a figure out of a screenshot (with grid preview mode)
 *   image-to-docx skill - the end-to-end transcription workflow and its pitfalls
 */
import { existsSync, readdirSync, readFileSync, writeFileSync, unlinkSync } from 'node:fs'
import { dirname, isAbsolute, join, resolve } from 'node:path'
import { homedir, tmpdir } from 'node:os'
import { randomBytes } from 'node:crypto'

const HERE = dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1'))
const BUILDER = join(HERE, 'builder.py')
const CROP = join(HERE, 'crop.py')

let cachedPython

function bundledPythons() {
  const found = []
  const push = (p) => {
    if (p && !found.includes(p) && existsSync(p)) found.push(p)
  }
  push(process.env.DSH_PYTHON)
  const home = process.env.DSH_HOME
  const roots = [home, join(homedir(), '.dsh')].filter(Boolean)
  for (const root of roots) {
    const runtimes = join(root, 'dsh-runtimes')
    try {
      for (const name of readdirSync(runtimes)) {
        push(join(runtimes, name, 'dependencies', 'python', 'python.exe'))
        push(join(runtimes, name, 'dependencies', 'python', 'bin', 'python3'))
      }
    } catch {
      /* no bundled runtime here */
    }
  }
  return found
}

async function resolvePython(ctx) {
  if (cachedPython) return cachedPython
  const bundled = bundledPythons()
  if (bundled.length) {
    cachedPython = bundled[0]
    return cachedPython
  }
  for (const name of ['python', 'python3', 'py']) {
    try {
      cachedPython = await ctx.subprocess.resolveExecutable(name)
      return cachedPython
    } catch {
      /* try the next name */
    }
  }
  throw new Error(
    'python not found. Install Python with python-docx + Pillow, or set DSH_PYTHON to the interpreter path.',
  )
}

function abs(baseDir, p) {
  return isAbsolute(p) ? p : resolve(baseDir, p)
}

async function runPython(ctx, script, specPath, signal) {
  const python = await resolvePython(ctx)
  const handle = ctx.subprocess.spawn({
    argv: [python, script, specPath],
    cwd: dirname(script),
    stdio: {
      stdin: 'ignore',
      stdout: { maxBytes: 4 << 20 },
      stderr: { maxBytes: 1 << 20 },
    },
    graceMs: 10_000,
    signal,
  })
  const outcome = await handle.done
  const stdout = handle.collected.stdout ? handle.collected.stdout.readFrom(0).text : ''
  const stderr = handle.collected.stderr ? handle.collected.stderr.readFrom(0).text : ''
  if (outcome.exitCode !== 0) {
    throw new Error(
      `python failed (exit ${outcome.exitCode}).\n${stderr.trim() || stdout.trim()}`,
    )
  }
  try {
    return JSON.parse(stdout.trim())
  } catch {
    throw new Error(`unexpected python output:\n${stdout}\n${stderr}`)
  }
}

const text = (s) => [{ type: 'text', text: s }]

export const inject = ['tools', 'skills', 'subprocess']

export async function apply(ctx) {
  const baseFor = (args) => (args && typeof args.base_dir === 'string' ? resolve(args.base_dir) : process.cwd())

  ctx.effect(() =>
    ctx.tools.register({
      name: 'docx_typeset',
      description:
        'Build a Word (.docx) document from a JSON layout spec using python-docx. ' +
        'Use this to re-typeset a transcribed page: blocks of text with {d}着重号{/d} and {b}bold{/b} markup, ' +
        'tab-stop columns (options A-D in 2 or 4 columns), shaded title banners, tables, cropped pictures, and explicit page breaks. ' +
        'It applies the hard-won details automatically: 着重号 uses w:em val="underDot" (dots BELOW the glyph) and those lines get at-least ' +
        'dot_min_line_pt (default 24) line spacing so the dots are not clipped; no "keep with next" is written (that would paint ■ squares in the Word margin) - use ' +
        '{"type":"pagebreak"} instead; tab stops are measured from the page margin. ' +
        'Relative picture paths resolve against base_dir. Returns the written path plus layout warnings (e.g. a column about to overflow its tab stop).',
      parameters: {
        type: 'object',
        properties: {
          spec: {
            type: 'object',
            description:
              'Layout spec. Shape: {out, page?{width_cm,height_cm,margin_cm{top,bottom,left,right,footer}}, ' +
              'body?{font,latin,size_pt,line_pt,dot_min_line_pt}, ' +
              'footer?{text,size_pt,align,color} (use {page}/{pages} for live PAGE/NUMPAGES fields), blocks:[...]}. ' +
              'Blocks: {"type":"title",runs:[{text,color,bold,size_pt,font,latin}],size_pt,line_pt,align,after_pt} | ' +
              '{"type":"text",text,indent_cm,first_line_cm,tabs_cm:[..],size_pt,line_pt,line_rule:"exact"|"at_least",before_pt,after_pt,align,fill,color,bold,font,latin} | ' +
              '{"type":"table",rows:[[cell,...]],widths_cm:[..],header,caption?} | ' +
              '{"type":"image",src,width_cm,align,caption,caption_size_pt} | ' +
              '{"type":"pagebreak"}. ' +
              'Per-block font/latin overrides the body font, which is how a page mixes 宋体 stems with 楷体 quoted material. ' +
              'Text markup: {d}着重号{/d}, {b}bold{/b}, {i}italic{/i}, {u}underline{/u}; a tab (real "\\t" or a literal backslash-t) starts a new tab-stop column. ' +
              'A leading "A.\\t" option label is normalized to "A. " automatically so it does not jump into the first column.',
            additionalProperties: true,
          },
          spec_path: { type: 'string', description: 'Read the spec JSON from this file instead of `spec`.' },
          out: { type: 'string', description: 'Output .docx path; overrides spec.out.' },
          base_dir: {
            type: 'string',
            description: 'Directory that relative paths (out, image src) resolve against. Defaults to the host process cwd, so pass absolute paths.',
          },
        },
        additionalProperties: false,
      },
      output: {
        schema: {
          type: 'object',
          properties: {
            path: { type: 'string' },
            counts: { type: 'object' },
            warnings: { type: 'array', items: { type: 'string' } },
          },
          additionalProperties: true,
        },
        render: (_args, value) => {
          const warnings = (value && value.warnings) || []
          const lines = [`Wrote ${value.path}`]
          if (warnings.length) {
            lines.push('', `Layout warnings (${warnings.length}):`)
            for (const w of warnings) lines.push(`- ${w}`)
          }
          return text(lines.join('\n'))
        },
      },
      isConcurrencySafe: () => false,
      async execute(args, exec) {
        const base = baseFor(args)
        const spec = args.spec_path
          ? JSON.parse(readFileSync(abs(base, args.spec_path), 'utf8'))
          : { ...(args.spec || {}) }
        if (args.out) spec.out = args.out
        if (!spec.out) throw new Error('spec needs an `out` path (or pass `out`)')
        spec.out = abs(base, spec.out)
        for (const block of spec.blocks || []) {
          if (block && block.type === 'image' && typeof block.src === 'string') {
            block.src = abs(base, block.src)
          }
        }
        const tmp = join(base, `.docx-spec-${randomBytes(4).toString('hex')}.json`)
        writeFileSync(tmp, JSON.stringify(spec, null, 2), 'utf8')
        try {
          return await runPython(ctx, BUILDER, tmp, exec.signal)
        } finally {
          try {
            unlinkSync(tmp)
          } catch {
            /* best effort */
          }
        }
      },
    }),
  )

  ctx.effect(() =>
    ctx.tools.register({
      name: 'docx_crop',
      description:
        'Crop a picture (figure, chart, diagram) out of a screenshot so it can be embedded with docx_typeset. ' +
        'Pass box=[x0,y0,x1,y1] in source pixels, or info=true to read the source size first, or grid=true to write a ' +
        'labelled coordinate grid preview that makes picking the box easy. Supports white-margin trimming and upscaling for print. ' +
        'For a PHOTOGRAPH of paper (not a clean screenshot) pass white_point=true so the grey sheet becomes pure white, ' +
        'otherwise the figure lands in Word sitting in a grey box; trim_threshold sets what counts as ink. ' +
        'For a question whose artwork is several panels (three rulers in a row, a picture beside its data) pass ' +
        'compose=true plus panels=[{box},...] and the panels are stitched into ONE picture, which is what docx_typeset ' +
        'needs because one image block holds one picture.',
      parameters: {
        type: 'object',
        properties: {
          src: { type: 'string', description: 'Source image (absolute, or relative to base_dir).' },
          out: { type: 'string', description: 'Output PNG path.' },
          box: {
            type: 'array',
            items: { type: 'number' },
            description: 'Crop rectangle [x0,y0,x1,y1] in source pixels.',
          },
          pad: { type: 'number', description: 'Extra pixels of margin around the box (default 0).' },
          trim: { type: 'boolean', description: 'Trim the margin left around the ink.' },
          trim_threshold: {
            type: 'number',
            description:
              'Pixels darker than this count as ink when trimming (default 245 for screenshots, 205 once white_point is set). ' +
              'Lower it to ignore faint pencil/stains; raise it to keep light-grey strokes.',
          },
          white_point: {
            type: ['number', 'boolean'],
            description:
              'Paper level mapped to pure white, e.g. 195. Pass true to auto-detect it from the histogram. ' +
              'Photographs of paper are grey (~200), so without this the crop keeps a grey background and the default 245 trim removes nothing.',
          },
          black_point: {
            type: 'number',
            description: 'Ink level mapped to black; only used together with a numeric white_point (auto mode picks it).',
          },
          grayscale: {
            type: 'boolean',
            description: 'Write an 8-bit grey PNG. Smaller and crisper for black-ink line art; implied by white_point.',
          },
          compose: {
            type: 'boolean',
            description: 'Stitch `panels` into a single picture instead of cropping one box.',
          },
          panels: {
            type: 'array',
            description:
              'Panels for compose mode. Each: {box:[x0,y0,x1,y1], src?, pad?, trim?}. `src` defaults to the top-level src, ' +
              'so several panels can come from one page.',
            items: {
              type: 'object',
              properties: {
                box: { type: 'array', items: { type: 'number' } },
                src: { type: 'string' },
                pad: { type: 'number' },
                trim: { type: 'boolean' },
              },
              additionalProperties: true,
            },
          },
          direction: { type: 'string', enum: ['horizontal', 'vertical'], description: 'Compose axis (default horizontal).' },
          align: {
            type: 'string',
            enum: ['bottom', 'top', 'center', 'left', 'right'],
            description: 'Cross-axis alignment of the panels. Bottom-align rows that share a baseline (e.g. rulers).',
          },
          gap: { type: 'number', description: 'Pixels between composed panels (default 32).' },
          scale: { type: 'number', description: 'Upscale factor for print quality (default 1).' },
          grid: { type: 'boolean', description: 'Draw a labelled grid instead of cropping.' },
          step: { type: 'number', description: 'Grid step in pixels (default 100).' },
          info: { type: 'boolean', description: 'Only report the source dimensions.' },
          base_dir: { type: 'string', description: 'Directory for relative paths (defaults to the host cwd).' },
        },
        required: ['src'],
        additionalProperties: false,
      },
      output: {
        schema: { type: 'object', properties: { path: { type: 'string' } }, additionalProperties: true },
        render: (_args, value) =>
          text(
            value.path
              ? `Wrote ${value.path} (${value.crop_width || value.width}x${value.crop_height || value.height}px)` +
                (value.panels ? `, ${value.panels} panels stitched` : '')
              : JSON.stringify(value),
          ),
      },
      isConcurrencySafe: () => true,
      async execute(args, exec) {
        const base = baseFor(args)
        const spec = { ...args, src: abs(base, args.src) }
        delete spec.base_dir
        if (Array.isArray(spec.panels)) {
          spec.panels = spec.panels.map((panel) =>
            panel && typeof panel.src === 'string' ? { ...panel, src: abs(base, panel.src) } : panel,
          )
        }
        if (!spec.out) spec.out = 'figure.png'
        spec.out = abs(base, spec.out)
        const tmp = join(tmpdir(), `docx-crop-${randomBytes(4).toString('hex')}.json`)
        writeFileSync(tmp, JSON.stringify(spec, null, 2), 'utf8')
        try {
          return await runPython(ctx, CROP, tmp, exec.signal)
        } finally {
          try {
            unlinkSync(tmp)
          } catch {
            /* best effort */
          }
        }
      },
    }),
  )

  ctx.effect(() =>
    ctx.skills.register({
      name: 'image-to-docx',
      description:
        '把截图/照片里的试卷、讲义、练习题识别成排版一致的 Word 文档：保留着重号、分栏选项、拼音、插图与表格。',
      whenToUse:
        '用户给来一张或多张截图/扫描件并要求“提取出来重新排版成 Word/文档”，尤其是带加点字（着重号）、选项分栏、插图或表格的题目页时使用。',
      source: 'runtime',
      invocation: { modelInvocable: true, userInvocable: true },
      content: SKILL_BODY,
    }),
  )
}

const SKILL_BODY = `# 截图 → Word 重排（image-to-docx）

把截图里的版式忠实还原成 .docx。核心是两步：**逐字转写**（人眼/模型读图）+ **用本插件的工具落地排版**。

## 流程

1. **先探图**：\`docx_crop info=true\` 读出每张图的像素尺寸；需要裁剪插图时先 \`grid=true\` 生成带坐标网格的预览图，再定裁剪框。
   定位插图有个不靠眼睛也稳的办法：先用 OCR（Windows 自带 \`Windows.Media.Ocr\`，zh-Hans-CN）拿到每行文字的 y 区间——**两行文字之间的“空带”就是插图**；再用墨迹掩膜（只取低饱和度的暗像素，红笔/蓝笔天然被排除）量出真实边界。
   - **拍照件必须先美白**：手机拍的纸底是 ~200 的灰，而 \`trim\` 的阈值是 245，一个像素都裁不掉，图会带整块灰底贴进 Word。加 \`white_point:true\`（自动）或 \`white_point:195, black_point:60\`，再配 \`trim_threshold:205\`。
   - **一题多联图**（一排三把尺子、左右两幅商品图）：\`image\` 块一个只能放一张，用 \`compose:true, panels:[{box},{box},{box}], align:"bottom"\` 拼成一张；共底线的图用 \`align:"bottom"\` 才不会错位。
   - **躲开手写**：裁框下边界要停在印刷内容之后。学生常在图下方写解题过程，y1 多留几十像素就会连手写一起带进图里——先看墨迹逐行分布再定 y1。
2. **逐图精读转写**：一张一张按块（题干 stem → 选项 rows → 说明/引文 → 插图/表格）转写。不要跳读、不要“合理猜测”用词。
   - **拼音按原样保留**。题目常考“读音有误”，故意印错（如 稻穗 huì、窗框 kuāng、囚犯似 sì、滋 zhī）；改成正确读音就把题做废了。只需保证声调符号本身正确。
   - **水印/遮挡**：按上下文补全，并在最终回复里逐条列出补全位置与内容，让用户复核。
   - 全角括号、书名号、引号、破折号、省略号（……）按原文照录。
3. **标注格式**（写进 spec 的文本里）：
   - 加点字：\`{d}屹{/d}立\`（一个 run 一个字最稳，Word 会给每个字画点）
   - 加粗：\`{b}（　）{/b}\`（题干括号空、标题、表头）
   - 列：用 \`\\t\` 分隔，并在块上给 \`tabs_cm\`
4. **写 spec 并生成**：\`docx_typeset\`（传 \`spec\` 对象或 \`spec_path\`；路径尽量用绝对路径，否则给 \`base_dir\`）。
5. **目视验收**：Word COM 导出 PDF 再转 PNG，逐页比对（见下“验证”）。交付前清理临时文件。

## spec 速查

\`\`\`json
{
  "out": "D:/out/试卷.docx",
  "page": { "width_cm": 21, "height_cm": 29.7,
            "margin_cm": { "top": 2.2, "bottom": 2.2, "left": 2.2, "right": 2.2 } },
  "body": { "font": "宋体", "latin": "Times New Roman", "size_pt": 12, "line_pt": 20 },
  "blocks": [
    { "type": "title", "size_pt": 26, "line_pt": 34, "runs": [
        { "text": "26 秋四年级上册语文 ", "color": "000000" },
        { "text": "1-4 单元", "color": "0000E6" },
        { "text": "期中复习", "color": "FA0000" } ] },
    { "type": "text", "text": "专题三：选择题综合专项　　50道", "size_pt": 15, "bold": true,
      "align": "center", "fill": "FFE591", "line_pt": 24, "before_pt": 3, "after_pt": 9 },

    { "type": "text", "text": "1.\\t下列加点字读音全部正确的一项是{b}（　）{/b}",
      "tabs_cm": [0.78], "before_pt": 7 },
    { "type": "text", "text": "A.\\t屹{d}立{/d}（yì）\\t霎{d}时{/d}（shà）\\t鹅{d}卵{/d}（luǎn）石\\t崩裂（bēng）",
      "tabs_cm": [6.05, 9.4, 12.55], "indent_cm": 1.1 },

    { "type": "table", "rows": [["题号", "答案"], ["1", "A"]],
      "widths_cm": [3, 12], "header": true, "cell_align": "center" },
    { "type": "image", "src": "D:/crops/fig1.png", "width_cm": 8, "align": "center",
      "caption": "图1 爬山虎的脚", "caption_size_pt": 10.5 },
    { "type": "pagebreak" }
  ]
}
\`\`\`

## 踩过的坑（本插件已内置规避，勿重复犯错）

| 坑 | 现象 | 正确做法 |
|---|---|---|
| **着重号画在字上** | \`w:em val="dot"\` 会把点画在字顶 | 必须 \`val="underDot"\`；也不要用“下划线”冒充 |
| **着重号被裁一半** | 行距“固定值”时 Word 把字下方的点切掉 | 带着重号的行 ≥ \`body.dot_min_line_pt\`（默认 24pt）并用 \`line_rule:"at_least"\`；字号小（≤13pt）时可把该值降到 21–22pt，既保点又不撑高版面 |
| **整篇字号拍脑袋** | 用 15pt 排出来每题都多折一行，页数从 9 变 13 | 按截图反推：**行宽 cm ÷ 该行字数 = 单字宽**，再查 1pt=0.0353cm；同时用“分页是否与原图一致”当验收标准 |
| **引文与题干同字体** | 材料/句子本是楷体，排成宋体后一眼假 | \`{"type":"text","font":"楷体","first_line_cm":1.06}\` —— 段落级 \`font\`/\`latin\` 覆盖 \`body\` |
| **页脚页码写死** | 复制成“第 1 页 共 27 页”，合并后页数全错 | \`footer:{"text":"第 {page} 页 共 {pages} 页"}\`，插件写的是 PAGE/NUMPAGES 域 |
| **左侧一排 ■** | 给段落设“与下段同页/段前分页/段中不分页”后 Word 在页边显示黑方块 | 一律不用 keep-with-next；要控制在第几页开始，就在**上一段末尾**插入分页符（\`{"type":"pagebreak"}\`） |
| **选项整行跳到右边** | 行内标签 \`A.\\t\` 只替换了第一个制表符，剩下的跳到默认制表位 | 每个块显式给 \`tabs_cm\`；标签后的 \`\\t\` 与列 \`\\t\` 用同一套制表位（插件会估算列宽并告警“会溢出到默认制表位”） |
| **制表位叠加缩进** | 以为 \`tabs_cm\` 从首行缩进算起 | 制表位从**页面左边距**算起，不要再加 \`indent_cm\` |
| **LibreOffice 渲染崩溃** | 本机 LO Kit 报 0xC0000005 | 用 Word COM 导 PDF 再栅格化（见“验证”）；不要反复重试 LO |
| **仅看 PDF 验收** | PDF 里看着没问题，Word 里点被裁、页边有 ■ | 关键项要在 Word 里看；PDF 只验证分页与对齐 |
| **插图直接贴原截图** | 整页图糊、占满版心 | 先用 \`docx_crop\` 按图注区域裁出图（\`trim\` 去白边、\`scale\` 2x 提升印刷清晰度），再按版心比例设 \`width_cm\` |
| **拍照件的图带灰底** | 图贴进 Word 是一块灰方块，和白纸对不上；\`trim\` 也裁不掉任何边 | 纸底 ~200 灰，245 阈值看不见它。加 \`white_point:true\`（或 \`white_point:195, black_point:60\`）+ \`trim_threshold:205\`；输出会自动转灰度，四角必须是 255 才算对 |
| **一题多联图被截成一张** | 一排三把尺子只贴出一把，或选项图整体缺失 | \`compose:true, panels:[{box},…]\`，横向用 \`align:"bottom"\` 共底线，再按整体宽度设 \`width_cm\` |
| **手写混进插图** | 图里出现学生的解题笔迹，且与印刷内容黏在一起分不开 | \`trim_threshold\` 抬高（如 115→205 之间取值）只按重墨收紧边界；更可靠的是看墨迹逐行分布，把 y1 停在印刷内容结束处再裁 |
| **表格手写空格对齐** | 列不对齐、跨页表头丢失 | 用 \`{"type":"table"}\`：\`widths_cm\` 定列宽、\`header:true\` 自动重复表头、\`Table Grid\` 边框 |
| **同一 subagent 复制粘贴长文** | 长 spec 被截断 | spec 写到文件，用 \`spec_path\` 传入 |

## 验证（Windows / Word 已安装）

\`\`\`powershell
# docx -> pdf（Word COM）
$w = New-Object -ComObject Word.Application; $w.Visible = $false
$d = $w.Documents.Open($docx, $false, $true)
$d.ExportAsFixedFormat($pdf, 17); $d.Close(0); $w.Quit()
\`\`\`

再用 \`Windows.Data.Pdf\` 逐页 \`RenderToStreamAsync\`（DestinationWidth 1240）出 PNG，逐页比对：
标题配色、黄色色带、分栏列位、着重号是否在字下、分页是否与原图一致、有无孤行或半截行。
注意：PowerShell 5.1 把 .ps1 当 ANSI 读，脚本路径与内容保持纯 ASCII，中文路径用参数传。

## 交付

- 明确说明**只做了截图里出现的部分**（第几张图没有、哪些题只有题干），不要假装补全不存在的题。
- 列出水印/遮挡处的补全清单。
- \`present\` 交付 .docx；清理临时 PDF/PNG/中间 spec。
`