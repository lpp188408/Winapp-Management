# Winapp Management

一个轻量 Windows 小工具，用来查看当前打开的软件窗口、Explorer 文件夹，以及从常见窗口标题中识别出的文件名，并支持发送正常关闭请求。

当前版本：`v1.0.6`

版本记录见：[CHANGELOG.md](CHANGELOG.md)。

## 普通用户怎么安装

如果你不是程序员，先看这份说明：[Windows 安装和使用说明](INSTALL_WINDOWS.md)。

最终使用时只需要一个文件：

```text
WinappManagement.exe
```

把它复制到 Windows 电脑后双击运行即可。

如果你现在只有源码文件夹，在 Windows 上双击 `build-windows.bat` 可以生成这个 exe。

## 功能

- 三页签查看：应用程序、Office 文件、收藏。
- 查看可见应用窗口、进程名、PID、窗口标题和程序路径。
- 识别当前打开的 Explorer 文件夹路径。
- 从 Word、Excel、PowerPoint、PDF、记事本等常见标题格式中提取文件名；VS Code、Obsidian 和浏览器只按应用显示。
- 支持搜索、手动刷新和多选关闭。
- Office 文件单独页签显示，并读取 Windows 系统文件图标。
- Word、Excel、PowerPoint 会优先通过 Office COM 读取当前打开文件的真实目录；读取不到时才显示未识别目录。
- 每一行都支持单独“关闭”，发送 Windows 正常关闭消息，不强制结束进程。
- 点击路径可以切换到对应的已打开窗口。
- 支持收藏应用、文件夹和可识别真实路径的文件，并在收藏页按类型分栏。
- 软件窗口处于激活状态时每 4 秒自动刷新一次；窗口失去焦点或最小化后暂停自动刷新，降低后台占用。
- 连续点击刷新或关闭时会防止重复刷新叠加，减少 Windows `未响应`。
- 不写注册表，不安装服务，不要求管理员权限。

## 不包含

- 不扫描进程打开的完整文件句柄。
- 不默认强制结束进程。
- 不包含驱动、系统服务或后台常驻安装组件。
- 如果某个文件只能识别到文件名、拿不到真实文件路径，会提示无法收藏为可打开项。

## 开发运行

需要 Windows 和 .NET 8 SDK。

```powershell
dotnet run --project .\src\WinappManagement\WinappManagement.csproj
```

## 打包单文件 exe

```powershell
.\scripts\publish-windows.ps1
```

输出位置：

```text
dist\win-x64\WinappManagement.exe
```

这个 exe 是 self-contained 单文件发布，可复制到未安装 .NET SDK 的 Windows 机器上运行。

## 打包轻量 exe

如果你的 Windows 电脑已经安装 Microsoft .NET 8 Desktop Runtime，也可以生成 framework-dependent 轻量版：

```powershell
.\scripts\publish-windows-light.ps1
```

非程序员可以直接双击：

```text
build-windows-light.bat
```

输出位置：

```text
dist\win-x64-light\WinappManagement.exe
```

这个版本文件更小，但目标电脑需要先安装 .NET 8 Desktop Runtime。

## 验证指南

1. 打开几个普通软件，确认出现在“应用程序”页签并显示应用名、窗口标题、进程名和 PID。
2. 打开几个 Explorer 文件夹，确认出现在“应用程序”页签并显示路径。
3. 打开 Word、Excel、PowerPoint 文件，确认出现在“Office文件”页签。
4. 点击某一行右侧的“关闭”，确认原应用收到正常关闭请求；有未保存内容时应出现原应用保存提示。
5. 点击某一行的路径，确认对应窗口切换到前台。
6. 勾选多行后点击“关闭选中”，确认只发送正常关闭请求。
7. 点击星标收藏应用或文件夹，重启工具后确认收藏仍然存在，并在收藏页按“文件夹 / Office文件 / 应用程序”分栏显示。
8. 发布后将 `dist\win-x64\WinappManagement.exe` 复制到另一台 Windows 机器，直接双击运行。

## 说明

普通文件识别依赖窗口标题，不保证所有软件都能准确暴露当前文件名。Word、Excel、PowerPoint 会额外尝试读取 Office 当前打开文件路径；新建未保存、权限不一致或 Office 忙碌时仍可能无法识别目录。

收藏保存在当前 Windows 用户目录下：

```text
%APPDATA%\WinappManagement\favorites.json
```
