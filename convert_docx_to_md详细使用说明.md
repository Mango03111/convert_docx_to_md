# convert_docx_to_md.ps1 使用说明

## 简介

`convert_docx_to_md.ps1` 是一个用于将 `Word .docx` 文档转换为 `Markdown` 的 PowerShell 脚本。

它会直接读取 `docx` 内部结构，提取正文内容和内嵌图片，并在当前目录生成：

- 一个 Markdown 文档
- 一个单独的图片文件夹

该脚本适合图文教程、说明文档、普通资料整理这类场景，尤其适合将带图片的 Word 文档快速转为可继续编辑的 Markdown。

## 主要功能

- 将 `.docx` 文档转换为 `.md`
- 自动导出 Word 中的内嵌图片
- 将图片以 Markdown 图片链接形式插入文档
- 自动将正文中的 URL 转换为 Markdown 链接
- 识别 Word 标题并转换为 Markdown 标题
- 识别列表并转换为 Markdown 列表
- 尝试保留粗体和斜体格式
- 支持将 Word 表格转换为 Markdown 表格

## 适用范围

较适合：

- 普通段落正文
- 图文教程
- 带图片的说明文档
- 简单标题结构
- 简单列表
- 简单表格

不保证完全还原的内容：

- 非常复杂的表格布局
- 页眉页脚
- 脚注、批注、修订记录
- 复杂嵌套样式
- 非标准编号格式
- 特殊 Word 对象或嵌入内容

## 文件位置

脚本文件：

- [convert_docx_to_md.ps1](c:\Users\23900\Desktop\MNVT31\convert_docx_to_md.ps1)

建议将脚本放在待转换文档所在目录中执行，这样最方便。

## 参数说明

脚本支持以下参数：

- `-DocxPath`
  指定要转换的 `.docx` 文件名或路径。

- `-MarkdownPath`
  指定输出的 `.md` 文件名或路径。

- `-ImageDir`
  指定导出的图片文件夹名称或路径。

如果这三个参数都不传：

- 脚本会在当前目录查找 `.docx`
- 当前目录下必须且只能有一个 `.docx` 文件
- 如果不止一个，会报错

默认命名规则：

- Markdown 文件名：`原文件名.md`
- 图片目录名：`原文件名_images`

## 输出结果

执行成功后，通常会生成：

- `xxx.md`
- `xxx_images/`

例如：

- `教程.docx`

会生成：

- `教程.md`
- `教程_images/`

Markdown 中的图片会以相对路径引用，例如：

```md
![](教程_images/image1.jpeg)
```

## Windows 使用方法

### 运行环境

推荐环境：

- Windows PowerShell 5.1
- 或 PowerShell 7+

如果系统限制脚本执行，可以临时使用 `-ExecutionPolicy Bypass` 运行。

### 方式一：当前目录只有一个 docx

在脚本所在目录打开 PowerShell，执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1
```

适用于：

- 当前目录只有一个 `.docx`
- 希望直接按默认名称输出

### 方式二：手动指定输入输出

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1 `
  -DocxPath "图文教程.docx" `
  -MarkdownPath "图文教程.md" `
  -ImageDir "M图文教程_images"
```

适用于：

- 当前目录有多个 `.docx`
- 想自定义输出文件名
- 想自定义图片目录名

### 在 PowerShell 7 中运行

如果你使用的是 PowerShell 7，也可以执行：

```powershell
pwsh -File .\convert_docx_to_md.ps1 -DocxPath "你的文档.docx"
```

### Windows 使用建议

- 文档名和图片目录名尽量避免重复
- 如果图片目录已存在，脚本会删除并重新创建该目录
- 建议先备份已有输出目录，避免被覆盖

## Linux 使用方法

### 运行环境

Linux 不能直接使用 `powershell.exe`，需要安装：

- `PowerShell 7+`
- 命令通常为 `pwsh`

可先检查是否已安装：

```bash
pwsh --version
```

如果系统提示找不到 `pwsh`，说明需要先安装 PowerShell。

### 方式一：当前目录只有一个 docx

进入脚本所在目录后执行：

```bash
pwsh ./convert_docx_to_md.ps1
```

### 方式二：手动指定输入输出

```bash
pwsh ./convert_docx_to_md.ps1 \
  -DocxPath "example.docx" \
  -MarkdownPath "example.md" \
  -ImageDir "example_images"
```

### Linux 使用建议

- 建议使用 `pwsh`，不要用 `bash` 直接执行 `.ps1`
- 如果文件名包含中文，请确保终端和系统编码正常
- 建议先在小样本文档上测试一次输出效果
- 如需批量转换，建议另外封装批处理脚本

## 典型使用流程

### Windows

1. 将 `convert_docx_to_md.ps1` 放到 Word 文档所在目录
2. 打开 PowerShell
3. 切换到该目录
4. 执行脚本
5. 检查生成的 `.md` 文件和图片目录

示例：

```powershell
cd C:\Users\23900\Desktop\MNVT31
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1 -DocxPath "MNVT-31A图文教程.docx"
```

### Linux

1. 将脚本和 `.docx` 放到同一目录
2. 确认系统已安装 `pwsh`
3. 进入目录
4. 执行脚本
5. 检查生成的 `.md` 文件和图片目录

示例：

```bash
cd /path/to/your/files
pwsh ./convert_docx_to_md.ps1 -DocxPath "example.docx"
```

## 常见问题

### 1. 提示当前目录存在多个 docx

原因：

- 脚本在未传 `-DocxPath` 时，只允许当前目录有一个 `.docx`

解决方法：

- 手动指定 `-DocxPath`

例如：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\convert_docx_to_md.ps1 -DocxPath "目标文档.docx"
```

### 2. 图片目录被覆盖了

原因：

- 如果输出图片目录已存在，脚本会先删除再重建

解决方法：

- 先修改 `-ImageDir`
- 或先备份旧目录

### 3. Markdown 格式不完全符合预期

原因：

- Word 和 Markdown 的结构能力并不完全对应
- 复杂文档转换后通常仍需人工微调

建议：

- 转换后检查标题、列表、表格和图片位置
- 对复杂内容进行二次整理

### 4. Linux 无法运行

原因通常是：

- 没有安装 `pwsh`
- 不是用 `pwsh` 执行脚本

正确方式：

```bash
pwsh ./convert_docx_to_md.ps1
```

## 注意事项

- 本脚本只支持 `.docx`，不支持旧版 `.doc`
- 输出结果适合继续编辑，不保证与 Word 原样完全一致
- 复杂格式文档建议先做测试
- 批量使用前，建议先确认输出目录不会覆盖已有文件

## 总结

如果你的目标是：

- 从 Word 快速导出 Markdown
- 保留图文结构
- 后续继续整理、发布、上传图床

那么这个脚本是比较实用的工具。

推荐做法是：

- 先转换
- 再人工检查
- 最后根据发布平台做排版整理
