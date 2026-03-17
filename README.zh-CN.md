# convert_docx_to_md

一个用于将 `.docx` 文档转换为 Markdown 并自动导出图片的 PowerShell 脚本。

它适合图文教程、笔记、说明文档、发布稿等常见场景，尤其适合把带图片的 Word 文档快速转换为可继续编辑和发布的 Markdown。

[English README](./README.md)

## 功能特性

- 将 `.docx` 转换为 `.md`
- 自动导出 Word 内嵌图片到单独文件夹
- 在 Markdown 中自动插入图片引用
- 将正文中的普通 URL 转换为 Markdown 链接
- 将 Word 标题转换为 Markdown 标题
- 识别列表并转换为 Markdown 列表
- 尽量保留粗体和斜体
- 将简单 Word 表格转换为 Markdown 表格

## 环境要求

- Windows 下可使用 PowerShell 5.1 或更高版本
- 推荐使用 PowerShell 7+
- Linux 下需要安装 `pwsh`

该脚本仅支持 `.docx`，不支持旧版 `.doc`。

## 快速开始

如果当前目录下只有一个 `.docx` 文件，可以直接执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1
```

如果使用的是 PowerShell 7：

```powershell
pwsh -File .\convert_docx_to_md.ps1
```

## 使用方法

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

## 参数说明

- `-DocxPath`
  输入 `.docx` 文件的路径或文件名

- `-MarkdownPath`
  输出 Markdown 文件的路径或文件名

- `-ImageDir`
  导出图片目录的路径或目录名

如果不传参数：

- 脚本会在当前目录查找 `.docx`
- 当前目录必须且只能存在一个 `.docx`
- 输出文件默认命名为：
  - `原文件名.md`
  - `原文件名_images`

## 输出结果

例如转换：

```text
guide.docx
```

会生成：

```text
guide.md
guide_images/
```

Markdown 中的图片引用格式类似：

```md
![](guide_images/image1.jpeg)
```

## 推荐使用流程

1. 将 `convert_docx_to_md.ps1` 放到 `.docx` 文件所在目录。
2. 执行脚本。
3. 检查生成的 Markdown 内容。
4. 如果原文档结构较复杂，再对标题、表格、列表做人工微调。
5. 继续发布、整理或上传图床。

## 注意事项

- 如果图片输出目录已存在，脚本会先删除再重建。
- 复杂的 Word 格式转换后通常仍需要人工检查。
- 表格转换更适合简单表格。
- 复杂编号、复杂布局、批注、页眉页脚、修订记录等内容不保证完整保留。

## 局限性

这个脚本的目标是“实用转换”，不是“完美还原 Word 排版”。

以下内容通常无法完全保留：

- 页眉页脚
- 批注和修订记录
- 脚注
- 复杂版式结构
- 所有 Word 专有样式行为

## 示例

```powershell
cd C:\path\to\your\files
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1 -DocxPath "input.docx"
```

## 相关文件

- [convert_docx_to_md.ps1](./convert_docx_to_md.ps1)
- [详细中文使用说明](./convert_docx_to_md使用说明.md)

