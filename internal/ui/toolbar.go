//go:build windows

package ui

import (
	"fmt"
	"math"
	"os"
	"path/filepath"
	"runtime"
	"syscall"
	"unsafe"

	"colorpaste/internal/palette"
)

const (
	cols = 4
	rows = 4

	trayUID = 1
)

// Pixel sizes below are specified at 96 DPI (100% scaling) and get
// multiplied by the system's actual DPI scale factor once at startup (see
// applyDPIScale, called from RunToolbarPreview). This process is DPI-aware
// (see the SetProcessDPIAware call there, needed for an unrelated coordinate
// bug — see its comment), so Windows no longer bitmap-stretches our
// rendering to compensate for scaling on its own; scaling these ourselves is
// what keeps the toolbar the same physical size on a scaled display instead
// of shrinking down to its bare 96-DPI pixel dimensions.
var (
	cellSize  int32 = 26
	cellGap   int32 = 4
	panelPad  int32 = 4
	cornerRad int32 = 6

	// headerH and closeBtnW match the official Windows caption button size
	// (46x28 epx, measured directly off a real system close button — see
	// x.png) so the hand-drawn close button below is pixel-for-pixel the
	// same size/position/hover behavior as a real system close button —
	// native WS_CAPTION rendering can't give exactly one button of that
	// size (a plain caption always reserves a second, merely-disabled
	// button slot next to it; a small "tool window" caption omits that slot
	// but also shrinks the button below the official size).
	headerH   int32 = 28
	closeBtnW int32 = 46

	// dragMoveThreshold is how far (combined x+y pixels) the cursor must
	// move after a button-down on the grid before it counts as "dragging
	// the window" rather than "clicking a swatch".
	dragMoveThreshold int32 = 4

	marginTop int32 = 60
	marginRt  int32 = 40

	iconSize       int32 = 16
	iconMarginLeft int32 = 8
)

// applyDPIScale multiplies every base (96 DPI) pixel size above by scale,
// converting them in place to the current display's actual pixel sizes.
// Must run once, before the window is created or anything is measured.
func applyDPIScale(scale float64) {
	s := func(v int32) int32 { return int32(math.Round(float64(v) * scale)) }
	cellSize = s(cellSize)
	cellGap = s(cellGap)
	panelPad = s(panelPad)
	cornerRad = s(cornerRad)
	headerH = s(headerH)
	closeBtnW = s(closeBtnW)
	dragMoveThreshold = s(dragMoveThreshold)
	marginTop = s(marginTop)
	marginRt = s(marginRt)
	iconSize = s(iconSize)
	iconMarginLeft = s(iconMarginLeft)
	closeGlyphFontHeight = s(closeGlyphFontHeight)
}

var (
	className    = "ColorPasteToolbarWnd"
	selectedIdx  = 4 // default selection: row 2, swatch 1 (柠檬黄/yellow)
	hwndToolbar  syscall.Handle
	onSelectFunc func(palette.Color)
	appIcon      syscall.Handle

	hoverClose    bool
	trackingMouse bool

	// Grid drag-tracking state: any button-down on the grid (whether on a
	// swatch or a gap) can turn into a window move if the cursor travels far
	// enough before button-up; otherwise it's a plain color-select click.
	gridTracking    bool
	gridDragMoving  bool
	gridClickIdx    = -1
	dragStartCursor point
	dragStartWinPos point
)

// gridSize is the 4x4 swatch grid's own size (excludes the header).
func gridSize() (w, h int32) {
	w = panelPad*2 + int32(cols)*cellSize + int32(cols-1)*cellGap
	h = panelPad*2 + int32(rows)*cellSize + int32(rows-1)*cellGap
	return
}

// windowSize is the full window size: header strip + grid below it.
func windowSize() (w, h int32) {
	gw, gh := gridSize()
	return gw, headerH + gh
}

func cellRect(i int) rect {
	col := int32(i % cols)
	row := int32(i / cols)
	left := panelPad + col*(cellSize+cellGap)
	top := panelPad + row*(cellSize+cellGap)
	return rect{Left: left, Top: top, Right: left + cellSize, Bottom: top + cellSize}
}

func hitTest(x, y int32) int {
	for i := 0; i < cols*rows; i++ {
		r := cellRect(i)
		if x >= r.Left && x < r.Right && y >= r.Top && y < r.Bottom {
			return i
		}
	}
	return -1
}

func closeBtnRect() rect {
	w, _ := windowSize()
	return rect{Left: w - closeBtnW, Top: 0, Right: w, Bottom: headerH}
}

func inRect(x, y int32, r rect) bool {
	return x >= r.Left && x < r.Right && y >= r.Top && y < r.Bottom
}

func abs32(v int32) int32 {
	if v < 0 {
		return -v
	}
	return v
}

func selectColor(hwnd syscall.Handle, i int) {
	if i == selectedIdx {
		return
	}
	selectedIdx = i
	c := palette.Colors[i]
	fmt.Printf("[toolbar] selected #%d %s %s\n", i+1, c.Name, c.Hex)
	if onSelectFunc != nil {
		onSelectFunc(c)
	}
	invalidate(hwnd)
}

// dragWindow lets the OS move this borderless popup as if the user grabbed
// its title bar, for clicks on the header strip outside the close button.
func dragWindow(hwnd syscall.Handle) {
	procReleaseCapture.Call()
	procSendMessageW.Call(uintptr(hwnd), wmNCLButtonDown, htCaption, 0)
}

func armMouseLeaveTracking(hwnd syscall.Handle) {
	tme := trackMouseEventT{
		dwFlags:   tmeLeave,
		hwndTrack: hwnd,
	}
	tme.cbSize = uint32(unsafe.Sizeof(tme))
	procTrackMouseEvent.Call(uintptr(unsafe.Pointer(&tme)))
}

func setHoverClose(hwnd syscall.Handle, v bool) {
	if hoverClose != v {
		hoverClose = v
		invalidate(hwnd)
	}
}

func wndProc(hwndArg, msgArg, wparam, lparam uintptr) uintptr {
	hwnd := syscall.Handle(hwndArg)
	msg := uint32(msgArg)
	switch msg {
	case wmPaint:
		onPaint(hwnd)
		return 0
	case wmMouseMove:
		if !trackingMouse {
			armMouseLeaveTracking(hwnd)
			trackingMouse = true
		}
		x := loword(lparam)
		y := hiword(lparam)
		setHoverClose(hwnd, y < headerH && inRect(x, y, closeBtnRect()))

		if gridTracking {
			var cp point
			procGetCursorPos.Call(uintptr(unsafe.Pointer(&cp)))
			dx := cp.X - dragStartCursor.X
			dy := cp.Y - dragStartCursor.Y
			if !gridDragMoving && abs32(dx)+abs32(dy) >= dragMoveThreshold {
				gridDragMoving = true
			}
			if gridDragMoving {
				procSetWindowPos.Call(uintptr(hwnd), 0,
					uintptr(int32(dragStartWinPos.X+dx)), uintptr(int32(dragStartWinPos.Y+dy)),
					0, 0, swpNoSize|swpNoZorder|swpNoActivate)
			}
		}
		return 0
	case wmMouseLeave:
		trackingMouse = false
		setHoverClose(hwnd, false)
		return 0
	case wmLButtonDown:
		x := loword(lparam)
		y := hiword(lparam)
		if y < headerH {
			if inRect(x, y, closeBtnRect()) {
				procDestroyWindow.Call(uintptr(hwnd))
			} else {
				dragWindow(hwnd)
			}
			return 0
		}
		gridClickIdx = hitTest(x, y-headerH)
		gridTracking = true
		gridDragMoving = false
		procGetCursorPos.Call(uintptr(unsafe.Pointer(&dragStartCursor)))
		var wr rect
		procGetWindowRect.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&wr)))
		dragStartWinPos = point{wr.Left, wr.Top}
		procSetCapture.Call(uintptr(hwnd))
		return 0
	case wmLButtonUp:
		if gridTracking {
			gridTracking = false
			procReleaseCapture.Call()
			if !gridDragMoving && gridClickIdx >= 0 {
				selectColor(hwnd, gridClickIdx)
			}
		}
		return 0
	case wmKeyDown:
		if wparam == vkEscape {
			procDestroyWindow.Call(uintptr(hwnd))
		}
		return 0
	case wmInitTray:
		addTrayIcon(hwnd)
		return 0
	case wmTrayIcon:
		switch uint32(lparam) {
		case wmRButtonUp:
			showTrayMenu(hwnd)
		case wmLButtonUp:
			procShowWindow.Call(uintptr(hwnd), swShow)
			procSetForegroundWindow.Call(uintptr(hwnd))
		}
		return 0
	case wmClose:
		procDestroyWindow.Call(uintptr(hwnd))
		return 0
	case wmDestroy:
		removeTrayIcon(hwnd)
		procPostQuitMessage.Call(0)
		return 0
	}
	r, _, _ := procDefWindowProcW.Call(hwndArg, msgArg, wparam, lparam)
	return r
}

func invalidate(hwnd syscall.Handle) {
	procInvalidateRect.Call(uintptr(hwnd), 0, 0)
}

// CurrentColor returns the palette color currently selected in the toolbar.
func CurrentColor() palette.Color {
	return palette.Colors[selectedIdx]
}

// ToolbarWindow returns the toolbar's own HWND (as a raw handle value) so
// other packages (e.g. the global mouse hook) can recognize and ignore
// clicks that land on the toolbar itself.
func ToolbarWindow() syscall.Handle {
	return hwndToolbar
}

func showTrayMenu(hwnd syscall.Handle) {
	var pt point
	procGetCursorPos.Call(uintptr(unsafe.Pointer(&pt)))

	hMenu, _, _ := procCreatePopupMenu.Call()
	if hMenu == 0 {
		return
	}
	defer procDestroyMenu.Call(hMenu)

	procAppendMenuW.Call(hMenu, mfString, idmHide, uintptr(unsafe.Pointer(utf16ptr("隐藏"))))
	procAppendMenuW.Call(hMenu, mfString, idmExit, uintptr(unsafe.Pointer(utf16ptr("退出"))))

	procSetForegroundWindow.Call(uintptr(hwnd))
	cmd, _, _ := procTrackPopupMenu.Call(
		hMenu, tpmRightButton|tpmReturnCmd,
		uintptr(pt.X), uintptr(pt.Y),
		0, uintptr(hwnd), 0,
	)
	switch cmd {
	case idmHide:
		procShowWindow.Call(uintptr(hwnd), swHide)
	case idmExit:
		procDestroyWindow.Call(uintptr(hwnd))
	}
}

func addTrayIcon(hwnd syscall.Handle) {
	var nid notifyIconDataW
	nid.cbSize = uint32(unsafe.Sizeof(nid))
	nid.hWnd = hwnd
	nid.uID = trayUID
	nid.uFlags = nifMessage | nifIcon | nifTip
	nid.uCallbackMessage = wmTrayIcon
	nid.hIcon = appIcon
	putUTF16(nid.szTip[:], "ColorPaste")
	procShellNotifyIconW.Call(nimAdd, uintptr(unsafe.Pointer(&nid)))
}

func removeTrayIcon(hwnd syscall.Handle) {
	var nid notifyIconDataW
	nid.cbSize = uint32(unsafe.Sizeof(nid))
	nid.hWnd = hwnd
	nid.uID = trayUID
	procShellNotifyIconW.Call(nimDelete, uintptr(unsafe.Pointer(&nid)))
}

// loadAppIcon loads assets/icon.ico from next to the running executable.
// If it's missing, the app still runs fine without a custom icon.
func loadAppIcon() syscall.Handle {
	exe, err := os.Executable()
	if err != nil {
		return 0
	}
	path := filepath.Join(filepath.Dir(exe), "assets", "icon.ico")
	if _, err := os.Stat(path); err != nil {
		return 0
	}
	h, _, _ := procLoadImageW.Call(
		0,
		uintptr(unsafe.Pointer(utf16ptr(path))),
		imageIcon, 0, 0,
		lrLoadFromFile|lrDefaultSize,
	)
	return syscall.Handle(h)
}

// closeGlyphFontHeight is the close button's "X" character height in
// logical units (negative selects character height over cell height, per
// CreateFontW convention), scaled in place by applyDPIScale like the sizes
// above.
var closeGlyphFontHeight int32 = -14

// drawX draws a plain text "X" centered in r. Letting DrawTextW (DT_CENTER |
// DT_VCENTER) do the centering avoids the hand-drawn two-line cross this
// used to be, whose intersection point could read a pixel high or low.
func drawX(hdc syscall.Handle, r rect, colorRef uintptr) {
	font, _, _ := procCreateFontW.Call(
		uintptr(closeGlyphFontHeight), 0, 0, 0,
		fwNormal,
		0, 0, 0,
		ansiCharset, outDefaultPrecis, clipDefaultPrecis, defaultQuality, defaultPitch,
		uintptr(unsafe.Pointer(utf16ptr("Segoe UI"))),
	)
	oldFont, _, _ := procSelectObject.Call(uintptr(hdc), font)
	oldColor, _, _ := procSetTextColor.Call(uintptr(hdc), colorRef)
	oldBkMode, _, _ := procSetBkMode.Call(uintptr(hdc), transparentBkMode)

	rc := r
	procDrawTextW.Call(uintptr(hdc), uintptr(unsafe.Pointer(utf16ptr("X"))), ^uintptr(0),
		uintptr(unsafe.Pointer(&rc)), dtCenter|dtVCenter|dtSingleLine)

	procSetBkMode.Call(uintptr(hdc), oldBkMode)
	procSetTextColor.Call(uintptr(hdc), oldColor)
	procSelectObject.Call(uintptr(hdc), oldFont)
	procDeleteObject.Call(font)
}

func onPaint(hwnd syscall.Handle) {
	var ps paintStruct
	hdcRes, _, _ := procBeginPaint.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&ps)))
	hdc := syscall.Handle(hdcRes)

	w, _ := windowSize()
	_, gh := gridSize()

	// Header: flat Windows-11-style caption strip with a single hand-drawn
	// close button matching the real system button's size/position exactly
	// (46 epx wide, full header height) and its hover behavior (solid
	// system close-red fill, white glyph) — see the headerH/closeBtnW
	// comment for why this isn't native WS_CAPTION rendering.
	headerBg, _, _ := procCreateSolidBrush.Call(rgb(0xf3, 0xf3, 0xf3))
	headerRect := rect{0, 0, w, headerH}
	procFillRect.Call(uintptr(hdc), uintptr(unsafe.Pointer(&headerRect)), headerBg)
	procDeleteObject.Call(headerBg)

	if appIcon != 0 {
		iconY := (headerH - iconSize) / 2
		procDrawIconEx.Call(uintptr(hdc), uintptr(iconMarginLeft), uintptr(iconY), uintptr(appIcon), uintptr(iconSize), uintptr(iconSize), 0, 0, diNormal)
	}

	// Colors measured directly off a real system close button (x.png):
	// #F3F3F3 idle background (already used above for the whole header),
	// #C42B1C hover fill, white glyph on hover.
	closeR := closeBtnRect()
	closeIconColor := rgb(0x5f, 0x5f, 0x5f)
	if hoverClose {
		hb, _, _ := procCreateSolidBrush.Call(rgb(0xc4, 0x2b, 0x1c))
		procFillRect.Call(uintptr(hdc), uintptr(unsafe.Pointer(&closeR)), hb)
		procDeleteObject.Call(hb)
		closeIconColor = rgb(0xff, 0xff, 0xff)
	}
	drawX(hdc, closeR, closeIconColor)

	// Swatch grid.
	gridBg, _, _ := procCreateSolidBrush.Call(rgb(0xf5, 0xf5, 0xf5))
	gridRect := rect{0, headerH, w, headerH + gh}
	procFillRect.Call(uintptr(hdc), uintptr(unsafe.Pointer(&gridRect)), gridBg)
	procDeleteObject.Call(gridBg)

	for i := 0; i < cols*rows; i++ {
		c := palette.Colors[i]
		r := cellRect(i)
		r.Top += headerH
		r.Bottom += headerH

		brush, _, _ := procCreateSolidBrush.Call(rgb(c.R, c.G, c.B))

		var pen uintptr
		if i == selectedIdx {
			pen, _, _ = procCreatePen.Call(psSolid, 3, rgb(0x33, 0x33, 0x33))
		} else {
			pen, _, _ = procCreatePen.Call(psSolid, 1, rgb(0xcc, 0xcc, 0xcc))
		}

		oldBrush, _, _ := procSelectObject.Call(uintptr(hdc), brush)
		oldPen, _, _ := procSelectObject.Call(uintptr(hdc), pen)

		procRoundRect.Call(uintptr(hdc), uintptr(r.Left), uintptr(r.Top), uintptr(r.Right), uintptr(r.Bottom), uintptr(cornerRad), uintptr(cornerRad))

		procSelectObject.Call(uintptr(hdc), oldBrush)
		procSelectObject.Call(uintptr(hdc), oldPen)
		procDeleteObject.Call(brush)
		procDeleteObject.Call(pen)
	}

	procEndPaint.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&ps)))
}

// RunToolbarPreview creates and shows the floating 4x4 color toolbar and
// blocks running the Win32 message loop until the window is closed (Esc key,
// the header close button, the tray icon's right-click "退出" menu item, or
// WM_CLOSE). The tray icon's right-click menu also has a "隐藏" item to hide
// the window without closing it; left-clicking the tray icon restores it.
// The header strip (other than the close button) can be dragged to move the
// window, and so can the grid itself — grabbing anywhere on it (a swatch or
// a gap) and moving the cursor drags the window, while releasing without
// much movement selects whatever swatch was under the cursor instead.
// onSelect, if non-nil, is invoked on every color change with the newly
// selected palette.Color.
//
// onReady, if non-nil, is invoked once, on this same OS thread, right after
// the window is created and shown but before the message loop starts — the
// right place for callers to install anything (like a low-level input hook)
// that must share this thread's message queue.
func RunToolbarPreview(onSelect func(palette.Color), onReady func()) error {
	runtime.LockOSThread()

	// Must happen before any coordinate-based call (WindowFromPoint,
	// GetCursorPos, CreateWindowExW, ...), in this package or internal/hook:
	// on a scaled display, a DPI-unaware process has its own such calls
	// silently rescaled by Windows to/from a virtualized 96-DPI coordinate
	// space, but the raw MSLLHOOKSTRUCT points the global mouse hook reads
	// are always physical pixels regardless of caller DPI awareness. Mixing
	// the two meant WindowFromPoint on a hook-reported point could be
	// rescaled a second time and land far outside the real physical screen
	// (observed: HWND 0 for any drag ending in roughly the right third of
	// this machine's 150%-scaled screen) — this call makes every Win32 call
	// in this process agree on physical coordinates.
	procSetProcessDPIAware.Call()

	// Now that Windows won't bitmap-stretch this window for us, scale our
	// own pixel sizes up to match, so the toolbar keeps its original
	// physical size on a scaled display instead of shrinking to its bare
	// 96-DPI dimensions.
	dpiRes, _, _ := procGetDpiForSystem.Call()
	if dpi := uint32(dpiRes); dpi > 0 {
		applyDPIScale(float64(dpi) / 96.0)
	}

	onSelectFunc = onSelect
	appIcon = loadAppIcon()

	hInstRes, _, _ := procGetModuleHandleW.Call(0)
	hInst := syscall.Handle(hInstRes)

	cursorRes, _, _ := procLoadCursorW.Call(0, idcArrow)

	cb := syscall.NewCallback(wndProc)

	wc := wndClassExW{
		cbSize:        uint32(unsafe.Sizeof(wndClassExW{})),
		style:         csHRedraw | csVRedraw,
		lpfnWndProc:   cb,
		hInstance:     hInst,
		hCursor:       syscall.Handle(cursorRes),
		hIcon:         appIcon,
		hIconSm:       appIcon,
		lpszClassName: utf16ptr(className),
	}
	atom, _, err := procRegisterClassExW.Call(uintptr(unsafe.Pointer(&wc)))
	if atom == 0 {
		return fmt.Errorf("RegisterClassExW failed: %w", err)
	}

	w, h := windowSize()
	screenW, _, _ := procGetSystemMetrics.Call(smCxScreen)
	x := int32(screenW) - w - marginRt
	y := int32(marginTop)

	hwndRes, _, err := procCreateWindowExW.Call(
		wsExTopMost|wsExToolWindow,
		uintptr(unsafe.Pointer(utf16ptr(className))),
		uintptr(unsafe.Pointer(utf16ptr("ColorPaste"))),
		wsPopup|wsVisible|wsBorder,
		uintptr(x), uintptr(y), uintptr(w), uintptr(h),
		0, 0, uintptr(hInst), 0,
	)
	if hwndRes == 0 {
		return fmt.Errorf("CreateWindowExW failed: %w", err)
	}
	hwndToolbar = syscall.Handle(hwndRes)

	procShowWindow.Call(uintptr(hwndToolbar), swShow)
	procUpdateWindow.Call(uintptr(hwndToolbar))
	procSetForegroundWindow.Call(uintptr(hwndToolbar))
	procPostMessageW.Call(uintptr(hwndToolbar), wmInitTray, 0, 0)

	if onReady != nil {
		onReady()
	}

	var m msgT
	for {
		ret, _, _ := procGetMessageW.Call(uintptr(unsafe.Pointer(&m)), 0, 0, 0)
		if int32(ret) <= 0 {
			break
		}
		procTranslateMessage.Call(uintptr(unsafe.Pointer(&m)))
		procDispatchMessageW.Call(uintptr(unsafe.Pointer(&m)))
	}
	return nil
}
