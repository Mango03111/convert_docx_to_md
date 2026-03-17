# convert_docx_to_md

A PowerShell script for converting `.docx` files into Markdown with extracted images.

It is designed for practical document conversion workflows such as tutorials, notes, guides, and image-heavy Word documents that need to be published or edited as Markdown.

[中文说明](./README.zh-CN.md)

## Features

- Convert `.docx` to `.md`
- Extract embedded images into a separate folder
- Insert extracted images as Markdown image links
- Convert plain URLs into Markdown links
- Convert Word headings into Markdown headings
- Recognize lists and render them as Markdown lists
- Preserve bold and italic formatting when possible
- Convert simple Word tables into Markdown tables

## Requirements

- PowerShell 5.1 or later on Windows
- PowerShell 7+ recommended for cross-platform use
- Linux users need `pwsh`

This script supports `.docx` only. Old `.doc` files are not supported.

## Quick Start

If the current directory contains exactly one `.docx` file:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1
```

Or with PowerShell 7:

```powershell
pwsh -File .\convert_docx_to_md.ps1
```

## Usage

### Windows

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1 `
  -DocxPath "input.docx" `
  -MarkdownPath "output.md" `
  -ImageDir "output_images"
```

### Linux

```bash
pwsh ./convert_docx_to_md.ps1 \
  -DocxPath "input.docx" \
  -MarkdownPath "output.md" \
  -ImageDir "output_images"
```

## Parameters

- `-DocxPath`
  Path or file name of the source `.docx`

- `-MarkdownPath`
  Path or file name of the generated Markdown file

- `-ImageDir`
  Output directory for extracted images

If no parameters are provided:

- the script searches the current directory for `.docx`
- there must be exactly one `.docx` file
- output names default to:
  - `original-name.md`
  - `original-name_images`

## Output

For example, converting:

```text
guide.docx
```

produces:

```text
guide.md
guide_images/
```

Image references inside the Markdown look like this:

```md
![](guide_images/image1.jpeg)
```

## Recommended Workflow

1. Put `convert_docx_to_md.ps1` in the same folder as the `.docx` file.
2. Run the script.
3. Review the generated Markdown.
4. Adjust headings, tables, or list formatting if the source document is complex.
5. Publish or continue editing.

## Notes

- Existing image output folders are deleted and recreated.
- Complex Word formatting may require manual cleanup after conversion.
- Table conversion works best for simple tables.
- Very complex numbering, layout, comments, headers, footers, and tracked changes are not fully preserved.

## Limitations

The script is intended for practical conversion, not perfect Word-to-Markdown fidelity.

It does not aim to fully preserve:

- headers and footers
- comments and revision history
- footnotes
- advanced layout structures
- all Word-specific styling behaviors

## Example

```powershell
cd C:\path\to\your\files
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1 -DocxPath "input.docx"
```

## Related Files

- [convert_docx_to_md.ps1](./convert_docx_to_md.ps1)
- [Chinese usage guide](./convert_docx_to_md使用说明.md)

## License

Add your preferred license here if you plan to publish the script publicly.
