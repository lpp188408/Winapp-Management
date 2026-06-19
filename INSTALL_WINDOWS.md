# Windows 安装和使用说明

这不是传统安装包。最终只需要一个 `WinappManagement.exe` 文件，复制到 Windows 电脑后双击即可运行。

## 最简单的使用方式

如果你已经拿到了 `WinappManagement.exe`：

1. 在 Windows 电脑上新建一个文件夹，例如：

   ```text
   D:\Tools\WinappManagement
   ```

2. 把 `WinappManagement.exe` 放进这个文件夹。
3. 双击 `WinappManagement.exe` 运行。
4. 如果 Windows 弹出安全提示，选择“更多信息”，再选择“仍要运行”。
5. 想放到桌面使用：右键 `WinappManagement.exe`，选择“发送到 > 桌面快捷方式”。

## 如果你现在只有源码文件夹

源码文件夹还不能直接双击运行。需要先在一台 Windows 电脑上生成 `WinappManagement.exe`。

只需要做一次：

1. 在 Windows 电脑上安装 `.NET 8 SDK`。
   - 下载地址：https://dotnet.microsoft.com/download/dotnet/8.0
   - 选择 Windows 的 SDK，不是 Runtime。
2. 把整个 `Winapp Management` 文件夹复制到 Windows 电脑。
3. 打开这个文件夹。
4. 双击 `build-windows.bat`。
5. 等待窗口提示“打包完成”。
6. 找到这个文件：

   ```text
   dist\win-x64\WinappManagement.exe
   ```

7. 以后只需要使用这个 `WinappManagement.exe`，不需要带源码文件夹。

如果没有生成 exe，请先看同一文件夹里的：

```text
build-log.txt
```

这个文件会记录失败原因。常见原因是没有安装 `.NET 8 SDK`，或只安装了 `.NET Runtime`。

如果双击后窗口一闪就没了，请确认你使用的是最新的 `build-windows.bat`。新版脚本会打开一个不会自动关闭的窗口，并在失败时显示 `build-log.txt` 的位置。

如果双击打包失败，也可以按下面方式手动打包：

1. 打开源码文件夹。
2. 按住 `Shift`，在文件夹空白处右键，选择“在终端中打开”或“在 PowerShell 中打开”。
3. 输入下面这行，然后按回车：

   ```powershell
   .\scripts\publish-windows.ps1
   ```

4. 等待完成后，找到这个文件：

   ```text
   dist\win-x64\WinappManagement.exe
   ```

## 推荐放置位置

可以放在这些地方之一：

```text
D:\Tools\WinappManagement\WinappManagement.exe
```

或：

```text
C:\Users\你的用户名\Apps\WinappManagement\WinappManagement.exe
```

不建议直接放在 `C:\Program Files`，因为那里经常需要管理员权限。

## 怎么确认它已经生效

1. 双击打开 `WinappManagement.exe`。
2. 再打开几个软件，比如记事本、浏览器、文件资源管理器。
3. 工具列表里应该能看到这些窗口。
4. 打开一个文件夹，工具里应该能看到“文件夹”分组。
5. 选中一个窗口，点“关闭选中”，对应软件应该收到正常关闭请求。

## 常见问题

### Windows 提示“不受信任”怎么办？

这是因为这个 exe 没有购买代码签名证书，不代表一定有问题。选择“更多信息 > 仍要运行”即可。

### 为什么不是安装包？

这个工具的设计目标就是轻量、免安装、便携。复制一个 exe 就能用，也方便删除。

### 怎么卸载？

关闭工具后，直接删除 `WinappManagement.exe` 所在文件夹即可。

### 是否需要管理员权限？

默认不需要。它不会安装服务，不会写注册表，也不会强制结束进程。
