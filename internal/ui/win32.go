//go:build windows

// Package ui implements the floating always-on-top color toolbar using raw
// Win32 API calls (no CGO, no third-party GUI library).
package ui

import (
	"syscall"
)

var (
	user32   = syscall.NewLazyDLL("user32.dll")
	gdi32    = syscall.NewLazyDLL("gdi32.dll")
	kernel32 = syscall.NewLazyDLL("kernel32.dll")
	shell32  = syscall.NewLazyDLL("shell32.dll")

	procRegisterClassExW    = user32.NewProc("RegisterClassExW")
	procCreateWindowExW     = user32.NewProc("CreateWindowExW")
	procDefWindowProcW      = user32.NewProc("DefWindowProcW")
	procShowWindow          = user32.NewProc("ShowWindow")
	procUpdateWindow        = user32.NewProc("UpdateWindow")
	procGetMessageW         = user32.NewProc("GetMessageW")
	procTranslateMessage    = user32.NewProc("TranslateMessage")
	procDispatchMessageW    = user32.NewProc("DispatchMessageW")
	procPostQuitMessage     = user32.NewProc("PostQuitMessage")
	procDestroyWindow       = user32.NewProc("DestroyWindow")
	procInvalidateRect      = user32.NewProc("InvalidateRect")
	procBeginPaint          = user32.NewProc("BeginPaint")
	procEndPaint            = user32.NewProc("EndPaint")
	procGetSystemMetrics    = user32.NewProc("GetSystemMetrics")
	procLoadCursorW         = user32.NewProc("LoadCursorW")
	procLoadIconW           = user32.NewProc("LoadIconW")
	procGetModuleHandleW    = kernel32.NewProc("GetModuleHandleW")
	procSetForegroundWindow = user32.NewProc("SetForegroundWindow")
	procFillRect            = user32.NewProc("FillRect")
	procReleaseCapture      = user32.NewProc("ReleaseCapture")
	procSendMessageW        = user32.NewProc("SendMessageW")
	procPostMessageW        = user32.NewProc("PostMessageW")
	procGetCursorPos        = user32.NewProc("GetCursorPos")
	procCreatePopupMenu     = user32.NewProc("CreatePopupMenu")
	procAppendMenuW         = user32.NewProc("AppendMenuW")
	procTrackPopupMenu      = user32.NewProc("TrackPopupMenu")
	procDestroyMenu         = user32.NewProc("DestroyMenu")
	procTrackMouseEvent     = user32.NewProc("TrackMouseEvent")
	procDrawIconEx          = user32.NewProc("DrawIconEx")
	procLoadImageW          = user32.NewProc("LoadImageW")
	procGetWindowRect       = user32.NewProc("GetWindowRect")
	procSetCapture          = user32.NewProc("SetCapture")
	procSetWindowPos        = user32.NewProc("SetWindowPos")
	procSetProcessDPIAware  = user32.NewProc("SetProcessDPIAware")
	procGetDpiForSystem     = user32.NewProc("GetDpiForSystem")
	procFindWindowW         = user32.NewProc("FindWindowW")
	procMessageBoxW         = user32.NewProc("MessageBoxW")

	procCreateMutexW = kernel32.NewProc("CreateMutexW")

	procCreateSolidBrush = gdi32.NewProc("CreateSolidBrush")
	procDeleteObject     = gdi32.NewProc("DeleteObject")
	procSelectObject     = gdi32.NewProc("SelectObject")
	procRoundRect        = gdi32.NewProc("RoundRect")
	procCreatePen        = gdi32.NewProc("CreatePen")
	procGetStockObject   = gdi32.NewProc("GetStockObject")
	procGetPixel         = gdi32.NewProc("GetPixel")
	procCreateFontW      = gdi32.NewProc("CreateFontW")
	procSetTextColor     = gdi32.NewProc("SetTextColor")
	procSetBkMode        = gdi32.NewProc("SetBkMode")

	procDrawTextW = user32.NewProc("DrawTextW")

	procShellNotifyIconW = shell32.NewProc("Shell_NotifyIconW")
)

const (
	wsPopup   = 0x80000000
	wsVisible = 0x10000000
	wsBorder  = 0x00800000

	wsExTopMost    = 0x00000008
	wsExToolWindow = 0x00000080

	csHRedraw = 0x0002
	csVRedraw = 0x0001

	swShow = 5

	wmDestroy       = 0x0002
	wmPaint         = 0x000F
	wmLButtonDown   = 0x0201
	wmLButtonUp     = 0x0202
	wmMouseMove     = 0x0200
	wmMouseLeave    = 0x02A3
	wmRButtonUp     = 0x0205
	wmKeyDown       = 0x0100
	wmClose         = 0x0010
	wmNCLButtonDown = 0x00A1
	wmApp           = 0x8000
	wmTrayIcon      = wmApp + 1
	wmInitTray      = wmApp + 2

	htCaption = 2

	vkEscape = 0x1B

	swpNoSize     = 0x0001
	swpNoZorder   = 0x0004
	swpNoActivate = 0x0010

	idcArrow       = 32512
	idiApplication = 32512

	psSolid   = 0
	nullBrush = 5

	fwNormal          = 400
	ansiCharset       = 0
	outDefaultPrecis  = 0
	clipDefaultPrecis = 0
	defaultQuality    = 0
	defaultPitch      = 0

	transparentBkMode = 1

	dtCenter     = 0x00000001
	dtVCenter    = 0x00000004
	dtSingleLine = 0x00000020

	smCxScreen = 0
	smCyScreen = 1

	swHide = 0

	tmeLeave = 0x00000002

	imageIcon      = 1
	lrLoadFromFile = 0x00000010
	lrDefaultSize  = 0x00000040

	diNormal = 0x0003

	nifMessage = 0x00000001
	nifIcon    = 0x00000002
	nifTip     = 0x00000004

	nimAdd    = 0x00000000
	nimModify = 0x00000001
	nimDelete = 0x00000002

	mfString       = 0x00000000
	tpmRightButton = 0x0002
	tpmReturnCmd   = 0x0100

	idmHide = 1000
	idmExit = 1001

	errorAlreadyExists = 183

	mbOK       = 0x00000000
	mbIconInfo = 0x00000040
	mbTopMost  = 0x00040000
	swRestore  = 9
)

type wndClassExW struct {
	cbSize        uint32
	style         uint32
	lpfnWndProc   uintptr
	cbClsExtra    int32
	cbWndExtra    int32
	hInstance     syscall.Handle
	hIcon         syscall.Handle
	hCursor       syscall.Handle
	hbrBackground syscall.Handle
	lpszMenuName  *uint16
	lpszClassName *uint16
	hIconSm       syscall.Handle
}

type point struct{ X, Y int32 }

type msgT struct {
	HWnd    syscall.Handle
	Message uint32
	WParam  uintptr
	LParam  uintptr
	Time    uint32
	Pt      point
}

type rect struct{ Left, Top, Right, Bottom int32 }

type paintStruct struct {
	Hdc         syscall.Handle
	FErase      int32
	RcPaint     rect
	FRestore    int32
	FIncUpdate  int32
	RgbReserved [32]byte
}

// notifyIconDataW mirrors the Win32 NOTIFYICONDATAW struct (only NIF_MESSAGE
// / NIF_ICON / NIF_TIP fields are actually populated by this app).
type notifyIconDataW struct {
	cbSize           uint32
	hWnd             syscall.Handle
	uID              uint32
	uFlags           uint32
	uCallbackMessage uint32
	hIcon            syscall.Handle
	szTip            [128]uint16
	dwState          uint32
	dwStateMask      uint32
	szInfo           [256]uint16
	uVersion         uint32
	szInfoTitle      [64]uint16
	dwInfoFlags      uint32
	guidItem         [16]byte
	hBalloonIcon     syscall.Handle
}

type trackMouseEventT struct {
	cbSize      uint32
	dwFlags     uint32
	hwndTrack   syscall.Handle
	dwHoverTime uint32
}

func putUTF16(dst []uint16, s string) {
	u, _ := syscall.UTF16FromString(s)
	copy(dst, u)
}

func loword(l uintptr) int32 { return int32(int16(uint16(l & 0xffff))) }
func hiword(l uintptr) int32 { return int32(int16(uint16((l >> 16) & 0xffff))) }

func rgb(r, g, b byte) uintptr {
	return uintptr(r) | uintptr(g)<<8 | uintptr(b)<<16
}

func utf16ptr(s string) *uint16 {
	p, _ := syscall.UTF16PtrFromString(s)
	return p
}
