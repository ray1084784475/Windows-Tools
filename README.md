# Windows实用工具集 | Windows Tools

**简体中文** | **English** Please scroll down

<a id="cn"></a>
## 简体中文
<div align="center">
	<img src="./image/ico.png" width=50 high=50>
	<h3>Windows实用工具集</h3>
	<img src="https://img.shields.io/badge/python-3.8%2B-blue">
	<img src="https://img.shields.io/badge/platform-Windows-lightgrey">
	<img src="https://img.shields.io/badge/License-MIT-yellow.svg">
</div>



### 简介

Windows实用工具集是一款专为 Windows 系统设计的实用工具集合，旨在提升用户的生产力和操作效率。它提供了窗口置顶、窗口半透明、音量控制和划词搜索等实用功能，全部可通过自定义快捷键快速访问。

### 功能特点

#### -窗口置顶
让任何窗口保持在最前面，方便参考文档、查看攻略或对比内容。
<img src="./image/1.gif">

#### -窗口半透明
将当前窗口设置为半透明，轻松查看下方窗口内容，无需切换窗口。
<img src="./image/2.gif">
<span style="color:green">更多玩法</span>：点击任务栏后按下快捷键即可实现任务栏透明。
<img src="./image/4.gif">

#### -划词搜索（<span style="color:blue;"><b>不完善，目前仅支持Windows自带应用</b></span>）
选中文本后一键搜索，支持自定义搜索引擎。
<span style="color:red"><b>默认快捷键会与记事本另存为快捷键相冲突，请更改。</b></span>
<img src="./image/3.gif">
#### -音量控制（<span style="color:red;"><b>不完善，不推荐使用</b></span>）
通过快捷键快速调节系统音量，特别适合台式机用户。
### 安装与使用
#### 从源码运行
1. 克隆或下载本项目
2. 安装依赖
3. 运行程序

#### 使用安装包
1. 下载最新版本的安装程序
2. 运行安装向导并按照提示完成安装
3. 程序将在系统托盘中运行

#### 打包为可执行文件
```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包为单文件可执行程序
pyinstaller -F -w --icon=ico.ico main.py

# 或者打包为文件夹形式
pyinstaller -w --icon=ico.ico main.py
```

### 自定义设置
您可以通过系统托盘菜单中的“设置”选项自定义所有功能的快捷键，保存或取消后请重新开启程序。
支持自定义搜索引擎URL，默认使用`Baidu`搜索。

### 系统要求
· Windows 10/11 （Windows 10未测试，不能确定能否实现功能）

· Python 3.8+ (如果从源码运行)

· 至少 100MB 可用磁盘空间

### 常见问题
Q: 程序启动后在哪里找到它？

A: 程序启动后会在系统托盘(通知区域)显示一个图标，右键点击可访问所有功能。

-
Q: 如何添加开机自启动？

A: 将软件创建快捷方式，并拖入启动文件夹`C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp`。此功能将在后续版本更新中加入。

-
Q: 划词搜索在某些应用程序中不起作用？

A: 某些应用程序使用自定义的文本控件，可能无法通过标准方法获取选中的文本。

### 贡献
欢迎提交问题报告和功能请求！如果您想贡献代码或提供翻译的英文代码，请发送邮件至`1084784475@qq.com`或`ray_eay@foxmail.com`。

### 支持

如果您遇到问题或有建议，请通过以下方式联系：

· 创建 GitHub Issue

· 发送邮件至: `1084784475@qq.com`或`ray_eay@foxmail.com`

### 捐赠

<img width=100 high=50 src="./image/Alipay.jpg">
<img width=150 high=150 src="./image/Alipay-Qrcode.jpg">

-


<img width=100 high=50 src="./image/Wechat-Payment.jpg">
<img width=150 high=150 src="./image/Wechat-Payment-Qrcode.jpg">
-

### 注意: 本软件是开源项目，仅供学习和交流使用。使用者应对自己的行为负责


- 

<a id="en"></a>
## English

Here's a non-human translation.Please contact `1084784475@qq.com` or `ray_eay@foxmail.com` if you have any errors or want to translate:

<div align="center">
	<img src="./image/ico.png" width=50 high=50>
	<h3>Windows Tools</h3>
	<img src="https://img.shields.io/badge/python-3.8%2B-blue">
	<img src="https://img.shields.io/badge/platform-Windows-lightgrey">
	<img src="https://img.shields.io/badge/License-MIT-yellow.svg">
</div>


### Introduction
Windows Tools is a utility suite specifically designed for Windows systems, aimed at enhancing user productivity and operational efficiency. It provides practical features including Window On Top, Window Transparency, Word Search, and Volume Control - all accessible via customizable hotkeys.

### Features
#### -Window On Top
Keep any window always on top for easy reference when viewing documents, guides, or comparing content.
<img src="./image/1.gif">

#### -Window Transparency
Set active windows to semi-transparent mode, allowing you to see underlying windows without switching applications.
<img src="./image/2.gif">
<span style="color:green">More gameplay</span> : Click the taskbar and press the shortcut key to make the taskbar transparent.
<img src="./image/4.gif">
#### -Word Search (<span style="color:blue;"><b>Limited: Currently only Windows native applications are supported</b></span>)
Search selected text instantly with your preferred search engine.
<span style="color:red"><b>The default shortcut will conflict with the save Notepad as shortcut. Please change it. </b></span>
<img src="./image/3.gif">
#### -Volume Control (<span style="color:red;"><b>Experimental: Not recommended</b></span>)
Adjust system volume rapidly using hotkeys - especially useful for desktop users.

### Installation & Usage
#### Run from Source
1. Clone or download the repository
2. Install dependencies
3. Execute the application

#### Installer Package
1. Download the latest installer
2. Run setup and follow installation prompts
3. The application will run in system tray

#### Build Executable

```bash
# Install PyInstaller
pip install pyinstaller

# Build single-file executable
pyinstaller -F -w --icon=ico.ico main.py

# Build folder distribution
pyinstaller -w --icon=ico.ico main.py
```

### Customization
You can customize the shortcut keys of all functions through the "Settings" option in the system tray menu. Please restart the program after saving or cancellations.Custom search engine URLs supported (Defaults to `Baidu`).

### System Requirements
· Windows 10/11 (Untested on Windows 10)  
· Python 3.8+ (For source execution)  
· Minimum 100MB available storage

### FAQs
Q: Where to find the application after launch? 
A: It runs in the system tray (notification area) - right-click the icon to access all features.

Q: How to enable auto-start?  
A: Create a shortcut in the Startup folder: `C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp` (To be implemented in future versions).

Q: Word Search doesn't work in some applications?  
A: Applications with custom text controls may prevent standard text selection detection.

### Contribution
Bug reports and feature requests are welcome! For code contributions or translations, contact: `1084784475@qq.com` or `ray_eay@foxmail.com`.

### Support
For assistance or suggestions:
· Open a GitHub Issue  
· Email: `1084784475@qq.com` or `ray_eay@foxmail.com`

### Donation

#### Alipay

<img width=100 high=50 src="./image/Alipay.jpg">
<img width=150 high=150 src="./image/Alipay-Qrcode.jpg">

-

#### Wechat Payment

<img width=100 high=50 src="./image/Wechat-Payment.jpg">
<img width=150 high=150 src="./image/Wechat-Payment-Qrcode.jpg">

-


### Note:  
This open-source software is intended for educational and communication purposes only. Users are solely responsible for their usage.
