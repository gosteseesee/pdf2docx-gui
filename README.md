# PDF 转 Word 工具

一款傻瓜式 PDF 转 Word 工具，拖入 PDF 文件即可自动转换为 Word 格式。

## 功能特点

- ✅ 拖拽文件即可转换
- ✅ 保持原 PDF 格式排版
- ✅ 自动选择输出目录
- ✅ 转换进度可视化
- ✅ 无需安装 Office 或 Adobe

## 使用方法

### 方法 1：下载 EXE（推荐）

1. 前往 [Releases](https://github.com/你的用户名/pdf2docx-gui/releases) 页面
2. 下载 `PDF转Word工具.exe`
3. 双击运行，无需安装

### 方法 2：从源码运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行
python pdf2docx_gui.py
```

## 打包 EXE（开发者）

本项目使用 GitHub Actions 自动打包，推送到 main 分支后自动生成 EXE。

手动打包：

```bash
# 安装打包工具
pip install pyinstaller

# 打包
pyinstaller --onefile --windowed --name "PDF转Word工具" pdf2docx_gui.py

# 生成的 EXE 在 dist/ 目录
```

## 技术栈

- **PDF 解析**: pdf2docx
- **GUI**: tkinter + tkinterdnd2
- **打包**: PyInstaller
- **CI/CD**: GitHub Actions

## 常见问题

**Q: 转换失败怎么办？**  
A: 确保 PDF 文件未加密，且文件大小 < 100MB

**Q: 格式保持不理想？**  
A: pdf2docx 对复杂排版支持有限，可尝试手动调整 PDF

**Q: 支持批量转换吗？**  
A: 当前版本仅支持单文件转换，批量功能开发中

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
