# Markdown2WordByTemplate

一组面向 Windows / Microsoft Word / MathType 的 Codex skills，用于把 Markdown 文档转换为套用 Word 模板的 `.docx`，并将文档中的 LaTeX 公式批量转换为可编辑的 MathType 对象。

这个项目适合以下场景：

- Markdown 是正文源文件；
- Word 模板已经定义好多级标题、正文、题注等样式(内置默认模板样式未调整，排版较丑)；
- 需要Word中的公式必须是 MathType `Equation.DSMT4` 对象，而不是 Word 原生 OMML 公式；
- Markdown 表格需要转换为三线表；
- Mermaid 流程图需要保留为 Visio 后处理占位；
- 图片和表格题注需要使用 Word `Caption` 样式和可更新编号域。

## 工作流

```text
input.md + template.docx
        │
        ▼
stage-1 DOCX
  - 套用 Word 模板
  - LaTeX 仍保留为源码
  - 表格转三线表
  - 图片/表格题注转 Word Caption
  - Mermaid 转 [VISIO:diagram-NN] 占位
        │
        ▼
final DOCX
  - LaTeX 转 MathType 对象
  - 行内公式保持 inline
  - 行间公式作为 display 公式
```

## Skills

本项目包含三个 skill：

| Skill | 作用 |
| --- | --- |
| `md-to-template-docx` | Markdown 转 stage-1 Word，保留 LaTeX 源码 |
| `word-mathtype-selective` | 将 Word 中的 `$...$` / `\[...\]` 转为 MathType 对象 |
| `md-template-mathtype-pipeline` | 编排前两个 skill，一条命令得到最终 Word |

建议保留三层结构，不要把所有逻辑硬合并成一个脚本。这样后续修复 Markdown 转换或 MathType 转换时，可以单独维护。

## 环境要求

### 基础环境

- Windows
- Microsoft Word
- MathType 6，且已激活
- Pandoc
- Python 3，需包含 `python-docx`、`lxml` 等依赖
- Codex skills 目录，通常是：

```text
%USERPROFILE%\.codex\skills
```

### MathType 要求

运行公式转换前必须：

1. 手动打开 Microsoft Word；
2. 新建或保留一个空白 Word 文档；
3. 确认 MathType Toggle TeX 可手动使用；
4. 关闭输入和输出 `.docx`；
5. 转换过程中不要操作 Word。

不要使用 WPS。Word 和 PowerShell 应保持相同权限级别，不要一个以管理员权限运行、另一个普通权限运行。

## 安装

直接从源码安装，把下面三个目录复制到 `$env:USERPROFILE\.codex\skills`：

```text
skills/md-to-template-docx
skills/word-mathtype-selective
skills/md-template-mathtype-pipeline
```

## 一条命令完整转换

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File "$env:USERPROFILE\.codex\skills\md-template-mathtype-pipeline\scripts\convert_md_to_mathtype_docx.ps1" `
  -MarkdownPath "C:\path\input.md" `
  -TemplatePath "C:\path\template.docx" `
  -OutputPath "C:\path\final.docx" `
  -Force
```

不传模板时使用内置默认模板：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File "$env:USERPROFILE\.codex\skills\md-template-mathtype-pipeline\scripts\convert_md_to_mathtype_docx.ps1" `
  -MarkdownPath "C:\path\input.md" `
  -OutputPath "C:\path\final.docx" `
  -Force
```

默认会生成：

```text
final_stage1_latex.docx
final.docx
final_visio_sources/
```

其中 `final_stage1_latex.docx` 是中间文件，公式仍是 LaTeX 源码，便于调试。

## 单独运行 stage-1

只生成套模板 Word，不转 MathType：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File "$env:USERPROFILE\.codex\skills\md-to-template-docx\scripts\run_md_to_template_docx.ps1" `
  -MarkdownPath "C:\path\input.md" `
  -TemplatePath "C:\path\template.docx" `
  -OutputPath "C:\path\stage1.docx"
```

## 单独运行 MathType 转换

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File "$env:USERPROFILE\.codex\skills\word-mathtype-selective\scripts\convert_word_latex_to_mathtype.ps1" `
  -InputPath "C:\path\stage1.docx" `
  -OutputPath "C:\path\final.docx" `
  -Force
```

## Markdown 输入规范

### 标题

Markdown 标题层级会直接映射到 Word 标题样式：

```markdown
# 一级标题
## 二级标题
### 三级标题
#### 四级标题
```

### 公式

行内公式：

```markdown
这是行内公式 $x_i \in \mathbb{R}^d$。
```

行间公式：

```markdown
\[
E = mc^2
\]
```

也支持 Markdown 中的 `$$...$$`，stage-1 会统一转成 `\[...\]`，再交给 MathType 转换。

### 表格

Markdown 表格会转换为三线表：

```markdown
| 指标 | 数值 |
| --- | --- |
| Precision | 0.95 |
```

注意：表格的正文映射为word模板中的正文未作单独处理，word模板中的正文样式设置了首行缩进2字符，表格中的正文同样会出现缩进2字符。

### Mermaid / Visio

Mermaid 代码块不会被 rasterize，而是替换为占位：

```markdown
```mermaid
flowchart TD
  A --> B
```
```

输出 Word 中会出现：

```text
[VISIO:diagram-01]
```

同时生成：

```text
final_visio_sources/diagram-01.mmd
final_visio_sources/manifest.json
```

后续可以用 Visio workflow 将占位替换为 `.vsdx` 矢量图。

## 图片和表格题注

题注识别发生在 stage-1，即 Pandoc 已经把 Markdown 转成 DOCX 之后。

### 图片题注识别优先级

1. 图片后方相邻段落，例如：

```markdown
![](./figures/framework.png)

图1 验证方法总体框架
```

2. Markdown 图片 alt 文本，例如：

```markdown
![验证方法总体框架](./figures/framework.png)
```

3. 如果没有题注，则插入占位：

```text
图N 请填写图题
```

当前版本不会自动把图片文件名 `framework.png` 当作题注。

如果使用 blockquote：

```markdown
![](./figures/framework.png)

> 图1 验证方法总体框架
```

### 表格题注

表格题注优先读取表格前方相邻段落：

```markdown
表1 检测指标对比

| 指标 | 数值 |
| --- | --- |
| AP | 0.91 |
```

如果缺失，则插入：

```text
表N 请填写表题
```

题注编号使用 Word 域：

```text
SEQ 图
SEQ 表
```

如果后续增删图表，在 Word 中按 `Ctrl+A` 后按 `F9` 更新编号。

## MathType 转换策略

`word-mathtype-selective` 不使用 Word 原生 OMML。它会：

- 连接到已打开的 Microsoft Word；
- 对 `$...$` 和 `\[...\]` 逐个定位；
- 调用 MathType Toggle TeX；
- 生成 `Equation.DSMT4` OLE 对象；
- 保留输入 DOCX 不变，另存输出 DOCX；
- 检查残留 LaTeX、MathType 对象数、Word OMath 对象数；
- 对行间公式段落做居中和空段清理。

### 已处理的兼容情况

- `$[200, 3000]$` 这类纯数值区间；
- `$c'$` 这类 prime 写法；
- 表格内 `$...$` 触发 MathType inline/display 弹窗；
- 行间公式后多余空段；
- 行间公式段落偏移或不居中。

## 常见问题

### PowerShell 提示没有打开 Word

先手动打开 Microsoft Word，保留一个空白文档，再运行脚本。

### MathType 授权弹窗阻塞

先在 Word 里手动插入/转换一次 MathType 公式，确认授权状态正常。

### 表格中的latex公式转换为MathType需要手动点击弹窗处理

目前表格正文中存在需要转换的latex公式时，会触发MathType弹窗bug，需要手动点击一下，点击之后会继续自动转换。

### Word 显示 `{SEQ 图 \* ARABIC}`

这是 Word 在显示域代码。按 `Alt + F9` 切换回显示编号结果。

## 目录结构

```text
work/skills/
  md-to-template-docx/
    SKILL.md
    assets/default-template.docx
    scripts/build_stage1_docx.py
    scripts/run_md_to_template_docx.ps1
    scripts/audit_stage1_docx.py
    scripts/preserve_math.lua

  word-mathtype-selective/
    SKILL.md
    scripts/convert_word_latex_to_mathtype.ps1

  md-template-mathtype-pipeline/
    SKILL.md
    scripts/convert_md_to_mathtype_docx.ps1

outputs/
  *.zip
  README-*.md
  test.md
```

## 设计原则

- stage-1 只负责生成结构正确的 Word，不创建 MathType 对象；
- stage-2 只负责公式转换，不重新处理 Markdown；
- pipeline skill 只负责编排，不复制底层逻辑；
- 尽量保留标准 LaTeX 源码，避免在 Markdown 阶段做 MathType 特定改写；
- 所有外部文档和模板保持不变，输出写入新文件。

## License

请根据你的发布需求补充 License。若无特殊限制，可考虑 MIT License。
