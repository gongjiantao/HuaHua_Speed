# 进大哥(华晨宇)房间快人一百步 (Screen Number Monitor)

[English Version Below](#english-version)

这是一个基于 Python 开发的屏幕区域监控小工具。它能像鹰眼一样盯着你指定的屏幕区域，一旦发现画面中出现**数字**，就会以闪电般的速度将其识别并**自动复制到剪贴板**，您只需要配合办公套件（例如：vivo办公套件、电脑手机互联相关软件），把电脑的复制内容同步到手机，即可快人一百步输入大哥的自建房间号。

**主要用途：** 抢房间号、快速提取验证码、监控动态数据变化等。

---

## ✨ 功能特点 (Features)

*   **🎯 精准选区**：点击按钮屏幕变暗，鼠标拖拽即可框选任意区域，所见即所得。
*   **👀 实时预览**：程序界面实时显示当前监控的画面，确保没选歪。
*   **🤖 智能识别**：基于 Google Tesseract OCR 引擎，精准识别画面中的所有数字。
*   **⚡ 极速复制**：识别到数字后毫秒级响应，自动写入剪贴板。
*   **🛡️ 抗干扰**：智能算法，自动过滤文字、符号，只提取纯数字（例如：`房间号：123456` -> 自动复制 `123456`）。
*   **🖥️ 高清适配**：完美支持 Windows 高分辨率屏幕（DPI 缩放），选区不漂移。

---

## 🚀 快速开始 (小白用户必读) / Quick Start

如果您不懂代码，只想要使用软件，请按照以下步骤操作：

### 第一步：安装必备的 OCR 引擎

本软件依赖 **Tesseract-OCR** 来识别文字，请先下载并安装它。

1.  **下载安装包** (任选一个下载)：
    - **官方下载 (v5.3.0)**：[点击下载](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.3.0.20221214.exe)
    - **SourceForge 镜像 (国内推荐)**：[点击下载](https://sourceforge.net/projects/tesseract-ocr.mirror/files/5.3.0/tesseract-ocr-w64-setup-v5.3.0.20221214.exe/download)
    - *注：v5.0 及以上版本均可，安装包约 50MB。*

2.  **安装步骤**：
    - 双击运行安装包。
    - **关键点**：建议保持默认安装路径 `C:\Program Files\Tesseract-OCR`。
    - *如果安装在其他盘，程序可能找不到它，需要您手动配置环境变量。*
    - 一路点击 Next 直到完成。

### 第二步：运行软件

1.  下载本项目的 `dist` 文件夹中的 `进大哥房间快人一百步.exe` (或者下载 Release 版本)。
2.  双击运行 `进大哥房间快人一百步.exe`。

---

## 📖 使用指南 (Usage)

1.  **启动软件**：打开软件后，您会看到一个简洁的操作界面。
2.  **选择区域**：
    - 点击 **【选择屏幕区域】(Select Screen Area)** 按钮。
    - 屏幕会蒙上一层灰色遮罩。
    - 按住鼠标左键，**框选** 您想要监控的数字区域（例如直播间右上角的房间号）。
    - 松开鼠标，软件会自动切回主界面。
3.  **确认预览**：
    - 看一下软件界面中间的 **[监控区域预览]**。
    - 确保预览图里能清晰看到您要监控的数字。
4.  **开始监控**：
    - 点击绿色 **【开始监控】(Start Monitoring)** 按钮。
    - 状态栏会变绿，显示“监控中...”。
5.  **自动复制**：
    - 现在您可以去干别的了。
    - 一旦那个区域出现了数字，软件下方的文本框会显示识别结果，并提示 **“已复制：xxxxxx”**。
    - 您只需要在配合办公套件（vivo办公套件），把电脑的复制内容同步到手机，即可快人一百步输入大哥房间号。
6.  **停止/退出**：
    - 点击红色 **【停止监控】(Stop Monitoring)** 按钮暂停。
    - 直接关闭窗口退出程序。

---

## 👨‍💻 开发者指南 (源码运行) / For Developers

如果您是开发者，想修改源码或自己编译：

### 1. 环境准备 (Prerequisites)
确保已安装 Python 3.8+。

```bash
# 克隆仓库 (Clone repo)
git clone https://github.com/your-username/your-repo-name.git
cd your-repo-name

# 安装依赖 (Install requirements)
pip install -r requirements.txt
```

### 2. 运行源码 (Run Source)
```bash
python screen_monitor.py
# 或者运行脚本 (Or run script)
./run_monitor.bat
```

### 3. 打包成 EXE (Build EXE)
如果您想自己生成 exe 文件：
```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name "进大哥房间快人一百步" screen_monitor.py
```
打包好的文件会在 `dist/` 目录下。

---

## ❓ 常见问题 (Q&A)

**Q: 软件打开报错 "Tesseract Not Found"？**
A: 说明您没有安装 Tesseract-OCR，或者没有安装在默认路径。请参考“快速开始”第一步重新安装，确保安装路径是 `C:\Program Files\Tesseract-OCR`。或者将 Tesseract 安装目录添加到系统的 PATH 环境变量中。

**Q: 为什么识别出来的数字不对？**
A: 请检查“监控区域预览”。
   - 选区是否太小？稍微框大一点。
   - 选区是否包含了太多杂乱文字？尽量只框住数字部分。
   - 背景是否太复杂？本软件对纯色背景上的黑色/白色数字识别效果最好。

**Q: 软件界面显示不全/按钮点不到？**
A: 这是因为您的系统开启了 DPI 缩放（如 150%）。最新版代码已修复此问题，请确保使用最新版本。

---

<a name="english-version"></a>
# Enter Big Brother (Hua Chenyu)'s Room One Step Ahead (Screen Number Monitor)

This is a Python-based screen area monitoring tool. It watches a specified area of your screen like a hawk, and once it detects **numbers**, it instantly recognizes them and **automatically copies them to the clipboard**. By using office suites (e.g., vivo Office Suite or PC-Phone connection software) to sync clipboard content from PC to mobile, you can enter Big Brother's room number one step ahead of others.

**Main Uses:** Grabbing room numbers, quickly extracting verification codes, monitoring dynamic data changes, etc.

## ✨ Features

*   **🎯 Precise Selection**: Click to dim the screen, drag to select any area. WYSIWYG.
*   **👀 Real-time Preview**: Shows the monitored area in real-time to ensure correct selection.
*   **🤖 Smart OCR**: Powered by Google Tesseract OCR engine for accurate digit recognition.
*   **⚡ Instant Copy**: Copies recognized digits to clipboard in milliseconds.
*   **🛡️ Noise Filtering**: Smart algorithm filters out text and symbols, extracting only pure digits (e.g., `Room: 123456` -> `123456`).
*   **🖥️ HD Support**: Fully supports Windows High DPI scaling.

## 🚀 Quick Start

### Step 1: Install OCR Engine (Required!)

This software depends on **Tesseract-OCR**.

1.  **Download Installer**:
    - **Official (v5.3.0)**: [Download Here](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v5.3.0.20221214.exe)
    - **SourceForge Mirror**: [Download Here](https://sourceforge.net/projects/tesseract-ocr.mirror/files/5.3.0/tesseract-ocr-w64-setup-v5.3.0.20221214.exe/download)

2.  **Install**:
    - Run the installer.
    - **Important**: Keep the default path `C:\Program Files\Tesseract-OCR`.
    - If installed elsewhere, please add the installation path to your system's PATH environment variable.

### Step 2: Run

1.  Download `进大哥房间快人一百步.exe` from the `dist` folder.
2.  Double-click to run.

---

## 📖 Usage Guide

1.  **Launch**: Open the software to see a simple interface.
2.  **Select Area**:
    - Click the **【Select Screen Area】** button.
    - The screen will dim.
    - Click and drag to **select** the area containing the digits you want to monitor (e.g., the room number).
    - Release the mouse, and it will return to the main interface.
3.  **Check Preview**:
    - Look at the **[Preview]** area in the middle.
    - Ensure the digits are clearly visible in the preview.
4.  **Start Monitoring**:
    - Click the green **【Start Monitoring】** button.
    - The status will turn green: "Monitoring...".
5.  **Auto Copy**:
    - You can now do other things.
    - Once digits appear in that area, the software will display the result below and show **"Copied: xxxxxx"**.
    - You only need to use an office suite (e.g., vivo Office Suite) to sync the PC clipboard to your phone, allowing you to enter Big Brother's room number one step ahead.
6.  **Stop/Exit**:
    - Click the red **【Stop Monitoring】** button to pause.
    - Close the window to exit.

---

**License**
MIT License
