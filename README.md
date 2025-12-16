# MDtoWORD-for-NJUST
一个能将Markdown格式的文件转换为word的小工具，word的排版默认为NJUST的学位论文排版。搭配一键将AI的输出内容转换为md文件的工具（deepshare)），有效解决AI输出内容复制粘贴到word格式不正确的难题。
# NJUST Thesis Formatter (Python 版)

**作者**: [[jimmyshuixin](https://github.com/jimmyshuixin "jimmyshuixin")]

**版本**: V3.2

这是一个基于 Python PyQt6 开发的论文格式转换工具，旨在将 Markdown 文件自动转换为符合 **南京理工大学 (NJUST)** 学位论文格式规范的 Word 文档 (`.docx`)。

## 🚀 功能特性

1. **双引擎支持**:
   - **Pandoc 引擎 (推荐)**: 完美支持数学公式 (`$E=mc^2$`)、复杂表格和参考文献。
   - **内置引擎**: 纯 Python 实现，无需安装 Pandoc 即可使用，适合简单的图文排版。

2. **拖拽转换**: 直接将 `.md` 文件拖入窗口即可生成 Word 文档。
3. **文件夹监控**: 选择一个文件夹，软件会自动监控，一旦有新的 `.md` 文件生成（例如由 Typora 导出或 AI 生成），会自动转换为 Word。
4. **自动排版**:
   - **字体**: 中文宋体，英文 Times New Roman (正文小四，标题加粗)。
   - **段落**: 正文固定行距 20 磅，首行缩进 2 字符。
   - **图片/表格**: 自动居中，图注/表注自动设置为五号字体。
   - **三线表**: 自动应用学术三线表样式。
   - **代码块**: 自动识别代码块并添加浅灰色背景，使用 Consolas 字体。

## 🛠️ 安装指南

### 1. 环境要求

- Python 3.8 或更高版本
- Windows / macOS / Linux

### 2. 安装依赖库

在项目目录下打开终端，运行以下命令：

```
pip install PyQt6 python-docx markdown beautifulsoup4 watchdog

```

> **注意**: `watchdog` 是用于文件夹监控的库，如果未安装，程序会自动降级为“轮询模式”，但建议安装以获得更好性能。

### 3. (可选) 安装 Pandoc

为了获得最佳的数学公式支持，建议安装 [Pandoc](https://pandoc.org/installing.html)。

- 程序会自动检测系统中的 Pandoc。
- 如果未检测到，会自动切换到内置 Python 引擎。

## 📖 使用方法

### 方式一：拖拽模式

1. 运行程序 `python main.py` (或双击打包好的 `exe`)。
2. 将写好的 Markdown 文件直接**拖入**程序窗口的虚线框内。
3. 程序会自动在 Markdown 同级目录下生成 `文件名_NJUST.docx`。

### 方式二：自动监控模式

1. 点击程序界面上的 **"模式二：选择并监控文件夹"** 按钮。
2. 选择你平时存放 Markdown 笔记的文件夹。
3. 当你在这个文件夹里**保存**或**新建**一个 `.md` 文件时，程序会自动检测并开始转换。

## 📦 如何打包为 EXE 可执行文件

如果你想把这个工具发给没有安装 Python 的同学使用，可以将其打包为独立的 `.exe` 程序。

### 步骤 1: 安装 PyInstaller

```
pip install pyinstaller

```

### 步骤 2: 执行打包命令

在项目根目录下打开终端，执行以下命令：

```
pyinstaller -F -w -n "NJUST论文助手" main.py

```

**参数解释**:

- `-F`: 打包成单个 exe 文件 (Onefile)。
- `-w`: 运行是不显示黑色控制台窗口 (Windowed)。
- `-n "..."`: 指定生成的文件名为 "NJUST论文助手.exe"。

### 步骤 3: 获取结果

打包完成后，打开项目目录下的 `dist` 文件夹，里面的 `NJUST论文助手.exe` 就是可以直接发送给别人的文件。

## 📝 常见问题 (FAQ)

**Q: 为什么生成的 Word 里公式是乱码？**
A: 请确保你安装了 **Pandoc**。如果没有 Pandoc，内置引擎无法完美渲染复杂的 LaTeX 公式。

**Q: 为什么图片没有显示？**
A: 请确保 Markdown 文件中的图片路径是正确的。如果是相对路径，请保证图片文件夹和 Markdown 文件在相对位置上没有变动。

**Q: 打开 exe 报错 "Failed to execute script..."**
A: 尝试去掉 `-w` 参数重新打包（`pyinstaller -F main.py`），然后在命令行运行生成的 exe，查看具体的错误报错信息（通常是因为缺少库）。

## 📜 开源协议

GNU General Public License v3.0
