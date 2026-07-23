# 开发日志

## 环境信息

- Go: go1.26.3 windows/amd64（已安装，`go` 命令全局可用）
- Git: 2.54.0（已安装），但 D:\Dev\ColorPaste 尚未 `git init`（当前非 git 仓库）
- 未检测到 gcc/mingw → UI 层决定使用纯 syscall 调用 user32/gdi32，不用 CGO，避免引入
  mingw 依赖（lxn/walk 等库也是纯 Go 实现，但优先用标准库 `syscall` 手写，减少依赖）
- 网络：go module 依赖首次拉取若失败，按 dev规范要求新开进程 10 秒后重试一次

## 任务节点状态

| # | 任务 | 状态 | 说明 |
|---|------|------|------|
| 1 | 需求开发文档 | ✅ 完成 | docs/需求开发文档.md |
| 2 | 项目骨架 (go.mod / 目录结构) | ✅ 完成 | cmd/colorpaste, internal/ui, internal/palette, internal/office, internal/hook |
| 3 | UI（悬浮工具栏） | ✅ 完成，待用户交互确认 | 26x26 色块/4px 间距；扁平 Win 风格标题栏（最小化 `—`、关闭 `×`，关闭悬停变红）；应用图标（4x4 迷你色板）；点击色块外任意区域可拖动整个窗口；托盘图标（右键"退出"，左键还原，最小化即隐藏到托盘） |
| 4 | 全局拖拽手势识别（WH_MOUSE_LL） | ✅ 完成 | internal/hook，见下方说明 |
| 5 | Word/Excel/PowerPoint COM 高亮自动化 | ✅ 首个可用版本 | internal/office，见下方说明；**尚未在用户真实 Word/Excel/PPT 里做过真实拖拽验证**（详见待确认事项） |
| 6 | 打包绿色版交付 | 部分完成 | bin/colorpaste.exe + bin/assets/icon.ico（图标需和 exe 放在同目录下的 assets 文件夹里） |

## 本轮实现说明

- **拖拽识别机制**：`internal/hook` 装的是全局低级鼠标钩子（WH_MOUSE_LL），只是"观察"，
  从不拦截/消费事件——目标软件（Word/Excel/PPT）会照常收到真实的鼠标事件，用户怎么拖，
  它就怎么原生选中文字，我们完全不用自己算选区。松开左键时，如果按下/松开两个点落在同一个
  窗口、且该窗口所属进程是 winword.exe/excel.exe/powerpnt.exe 之一、且横纵移动距离超过一点
  阈值，就用工具栏当前选中的颜色去调用 internal/office。
- **COM 自动化机制**：`internal/office` 只连接**已经在运行**的 Office 实例（`GetActiveObject`，
  绝不会自己启动 Word/Excel/PPT）；拿到目标 App 后直接读它自己的 `Selection`（即刚才用户真实
  拖拽产生的选区），分别用 Word 的 `Range.Shading.BackgroundPatternColor`、Excel 的
  `Selection.Interior.Color`、PowerPoint 的 `Selection.TextRange.Font.Highlight` 写入颜色。
  任何一步失败（软件没开 / 属性不存在 / 版本太旧不支持）都会被捕获并静默跳过，符合"不兼容则
  不生效"的需求，不会崩溃或误伤用户正常操作。
- COM 调用全部在一个专用、锁定 OS 线程、`CoInitialize` 过的 goroutine 里做（Office 的 COM
  对象是单线程套间 STA，必须固定同一线程访问）。
- 图标：`assets/icon.ico`（16/32/48/256 四个尺寸，PNG 压缩帧）用当前 16 色色板拼成 4x4 迷你图，
  由 `internal/ui` 在运行时通过 `LoadImageW` 从 exe 同目录的 `assets\icon.ico` 加载，同时用作
  窗口图标和托盘图标。

## 待处理 / 待用户确认事项

- **需要你亲自验证**（未在真实 Word 文档上测试——避免我的自动化脚本误改你正在编辑的真实文件）：
  1. 悬浮栏视觉：图标、最小化/关闭按钮悬停变红、26x26 网格是否满意；
  2. 拖动：点色块以外区域拖动是否顺滑；
  3. 最小化到托盘 / 托盘图标左键还原 / 右键退出是否正常；
  4. **核心功能**：打开一个已有 Word 文档，工具栏选中一个颜色，在 Word 里横向拖拽选中一段文字，
     松开后文字底纹是否变成了选中的颜色。（本机检测到有 WINWORD.EXE 正在运行，可以直接试）
  5. Excel/PowerPoint 同理（PowerPoint 的 `Font.Highlight` 只有 2019/365 及以上版本支持，旧版
     预期不生效）。
- 本机同时装了 WPS Office（进程名 wps/et 等）。当前只接的是微软 Office 的 ProgID
  （Word.Application/Excel.Application/PowerPoint.Application）——如果 WPS 注册了兼容这些
  ProgID（很多 WPS 安装会注册），有可能也能生效；没注册的话 WPS 场景会直接不生效，这也符合
  "不兼容则不生效"的预期，但如果你主要用 WPS 而不是微软 Office，需要另外适配 WPS 自己的
  COM 接口（ProgID 不同），需要你确认是否要做。

## 踩坑记录（新增）

## 有用信息 / 踩坑记录

- **Win32 WNDPROC 回调必须全部用 `uintptr` 参数**：用 `syscall.NewCallback` 包装 WndProc 时，
  如果参数里混用了 `uint32`（如 `msg uint32`）而不是全部 `uintptr`，栈上的参数偏移会错位，
  导致窗口"看起来创建成功"（RegisterClassExW / CreateWindowExW 都返回非 0），但消息循环会在
  某条消息上死锁，进程几秒后被系统标记为"未响应"。排查方法：`Get-Process | Responding` 变
  False，且对该窗口发同步消息（如 PrintWindow）会永久 hang。修复：WndProc 签名全部改成
  `func(hwnd, msg, wparam, lparam uintptr) uintptr`，内部再强转成需要的类型。
- **本机截图排查 DPI 坑**：本机是 150% 缩放（物理 2560x1600，逻辑约 1707x1067）。用非 DPI
  感知的 PowerShell 脚本（`System.Windows.Forms.Screen` + `CopyFromScreen`）截图时，我们自己
  这个非 DPI 感知的 Go 窗口即使真的画出来了、GetPixel 在窗口自己的 HDC 里验证颜色完全正确，
  截图里那块区域仍然可能拍不到内容（透出后面的窗口）。加一行 `SetProcessDPIAware()` 让截图
  脚本变成 DPI 感知后，用物理分辨率截全屏就能正常看到该窗口内容了。以后调试"Go 窗口画面对不
  对"优先用窗口自身 HDC 的 `GetPixel` 验证（不依赖截图/DPI），比截图更可靠。
- 未检测到 gcc/mingw，UI 层全部用标准库 `syscall` 直接调 user32.dll / gdi32.dll（无 CGO）。
  `GetModuleHandleW` 在 **kernel32.dll**，不是 user32.dll，容易挂错 DLL。
- **`Shell_NotifyIconW`（托盘图标）绝对不能在窗口消息循环 `GetMessage` 开始跑之前同步调用**：
  它需要跟 Explorer 走一遍 IPC 握手，如果调用它的时候本线程还没进入消息循环，会直接死等
  （进程零 CPU，`Get-Process` 里 `Responding` 变 False，且此后一条消息都收不到）。
  正确做法：`CreateWindowExW` 之后先 `PostMessage` 给自己一个自定义消息，等真正进了
  `GetMessage` 循环、这条消息被取出分发时，再在 `wndProc` 里调用 `Shell_NotifyIconW`。
- **`Get-Process ... | Responding` 在这台机器上对我们这种 `WS_EX_TOOLWINDOW` 无边框置顶弹窗
  不可靠**：多次观察到即使程序其实活得好好的（自己打的调试日志能连续证明消息循环仍在正常
  处理消息），`Responding` 也会读成 `False`，哪怕全程没做任何鼠标/键盘模拟。以后验证这类窗口
  是否卡死，优先看程序自己输出的日志有没有持续更新，而不是看这个属性。
- go-ole（`github.com/go-ole/go-ole`）首次 `go get` 一次性成功，未触发"需新开进程重试联网"
  的分支；如果以后遇到下载失败，按 dev规范第4条处理。
- Office COM 对象是 STA（单线程套间），必须固定在同一个 `runtime.LockOSThread()` 过、且
  `CoInitialize` 过的 goroutine 里访问，不能随便换 goroutine/线程调用。
