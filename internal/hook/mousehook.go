//go:build windows

// Package hook installs a global low-level mouse hook that detects a
// drag gesture (button down, move, button up) made outside the ColorPaste
// toolbar. It never blocks or alters the drag itself — the target
// application (Word/Excel/PowerPoint) sees and handles the real mouse
// events exactly as if ColorPaste weren't running, and ends up with its own
// native text selection. Once the button is released, if both endpoints
// were over the same supported app, this package asks internal/office to
// shade that selection with the toolbar's currently selected color.
package hook

import (
	"fmt"
	"syscall"
	"time"
	"unsafe"

	"colorpaste/internal/office"
	"colorpaste/internal/ui"
)

// ts returns a millisecond-precision timestamp for correlating log lines
// against exactly when the user did something.
func ts() string { return time.Now().Format("15:04:05.000") }

var (
	user32   = syscall.NewLazyDLL("user32.dll")
	kernel32 = syscall.NewLazyDLL("kernel32.dll")

	procSetWindowsHookExW   = user32.NewProc("SetWindowsHookExW")
	procUnhookWindowsHookEx = user32.NewProc("UnhookWindowsHookEx")
	procCallNextHookEx      = user32.NewProc("CallNextHookEx")
	procWindowFromPoint     = user32.NewProc("WindowFromPoint")
	procGetWindowThreadPid  = user32.NewProc("GetWindowThreadProcessId")

	procOpenProcess               = kernel32.NewProc("OpenProcess")
	procCloseHandle               = kernel32.NewProc("CloseHandle")
	procQueryFullProcessImageName = kernel32.NewProc("QueryFullProcessImageNameW")
)

const (
	whMouseLL = 14

	wmLButtonDown = 0x0201
	wmLButtonUp   = 0x0202

	hcAction = 0

	processQueryLimitedInformation = 0x1000

	// dragThreshold is the minimum combined pixel movement (dx+dy) between
	// button-down and button-up for a click to count as a drag gesture.
	dragThreshold = 8
)

type point struct{ X, Y int32 }

type msllhookstruct struct {
	pt          point
	mouseData   uint32
	flags       uint32
	time        uint32
	dwExtraInfo uintptr
}

var (
	hHook syscall.Handle

	dragActive    bool
	dragStartPt   point
	dragStartHwnd syscall.Handle
)

// mouseButtonEvent is the minimal data hookProc copies out before handing off
// to processEvents. Windows requires a low-level hook to return almost
// immediately (it gates *all* mouse input system-wide on that return); doing
// any of WindowFromPoint/OpenProcess/QueryFullProcessImageNameW inline here,
// as this package used to, meant one slow or wedged call — e.g. a target
// window that doesn't answer promptly — could stall mouse input for the
// entire desktop, not just this app. hookProc now only ever copies a struct
// and does a non-blocking channel send.
type mouseButtonEvent struct {
	down bool
	pt   point
}

var mouseEvents = make(chan mouseButtonEvent, 16)

func abs32(v int32) int32 {
	if v < 0 {
		return -v
	}
	return v
}

func windowFromPoint(pt point) syscall.Handle {
	h, _, _ := procWindowFromPoint.Call(uintptr(uint32(pt.X)) | uintptr(uint32(pt.Y))<<32)
	return syscall.Handle(h)
}

// exePathForWindow resolves the full image path of the process that owns
// hwnd, e.g. "C:\...\WINWORD.EXE".
func exePathForWindow(hwnd syscall.Handle) string {
	var pid uint32
	procGetWindowThreadPid.Call(uintptr(hwnd), uintptr(unsafe.Pointer(&pid)))
	if pid == 0 {
		return ""
	}
	h, _, _ := procOpenProcess.Call(processQueryLimitedInformation, 0, uintptr(pid))
	if h == 0 {
		return ""
	}
	defer procCloseHandle.Call(h)

	var buf [512]uint16
	size := uint32(len(buf))
	ret, _, _ := procQueryFullProcessImageName.Call(h, 0, uintptr(unsafe.Pointer(&buf[0])), uintptr(unsafe.Pointer(&size)))
	if ret == 0 {
		return ""
	}
	return syscall.UTF16ToString(buf[:size])
}

func hookProc(nCode, wparam, lparam uintptr) uintptr {
	if int32(nCode) == hcAction {
		data := (*msllhookstruct)(unsafe.Pointer(lparam))
		switch uint32(wparam) {
		case wmLButtonDown, wmLButtonUp:
			select {
			case mouseEvents <- mouseButtonEvent{down: uint32(wparam) == wmLButtonDown, pt: data.pt}:
			default:
				// processEvents is still catching up; dropping this one is
				// far better than blocking the hook to wait for room. Logged
				// from here would itself risk being the slow call the hook
				// can't afford, so it's counted and reported from
				// processEvents instead.
				droppedEvents++
			}
		}
	}
	r, _, _ := procCallNextHookEx.Call(0, nCode, wparam, lparam)
	return r
}

var droppedEvents int

// processEvents does all the actual work (WindowFromPoint, process lookups,
// the COM dispatch) off the hook thread, at its own pace.
func processEvents() {
	var seq int
	for ev := range mouseEvents {
		seq++
		if ev.down {
			dragStartPt = ev.pt
			dragStartHwnd = windowFromPoint(ev.pt)
			dragActive = true
			fmt.Printf("[hook] %s #%d down at (%d,%d) hwnd=%v dropped-so-far=%d\n",
				ts(), seq, ev.pt.X, ev.pt.Y, dragStartHwnd, droppedEvents)
		} else {
			fmt.Printf("[hook] %s #%d up at (%d,%d) dragActive=%v\n", ts(), seq, ev.pt.X, ev.pt.Y, dragActive)
			if dragActive {
				dragActive = false
				onDragEnd(ev.pt)
			}
		}
	}
}

func onDragEnd(endPt point) {
	if dragStartHwnd == 0 || dragStartHwnd == ui.ToolbarWindow() {
		fmt.Printf("[hook] %s onDragEnd bail: dragStartHwnd=%v (0 or toolbar)\n", ts(), dragStartHwnd)
		return
	}
	endHwnd := windowFromPoint(endPt)
	if endHwnd == 0 {
		fmt.Printf("[hook] %s onDragEnd bail: endHwnd=0\n", ts())
		return
	}

	// Compare app kind rather than requiring the exact same child HWND at
	// both ends: after a highlight is applied, Word/Excel/PowerPoint often
	// pop up their own floating "mini toolbar" right next to the selection,
	// which is a *different* child window of the same process. If the next
	// drag happens to end on/near it (very likely, since you naturally
	// reselect nearby text), a strict HWND match would wrongly reject the
	// gesture — which is exactly why only the very first highlight of a
	// session used to take effect.
	startExe := exePathForWindow(dragStartHwnd)
	startKind := office.DetectAppKind(startExe)

	dx := abs32(endPt.X - dragStartPt.X)
	dy := abs32(endPt.Y - dragStartPt.Y)
	// Excel/WPS Sheet select at cell granularity — a plain click (no drag at
	// all) already selects a whole, meaningful cell there, unlike Word/PPT
	// where a click with no movement just places a text cursor. So only
	// require real movement for the text-selection apps.
	requiresDrag := startKind != office.Excel && startKind != office.WPSSheet
	if requiresDrag && dx+dy < dragThreshold {
		fmt.Printf("[hook] %s onDragEnd bail: below threshold dx+dy=%d\n", ts(), dx+dy)
		return
	}

	var endKind office.AppKind
	if endHwnd == ui.ToolbarWindow() {
		// Our own toolbar is always-on-top, so a drag that ends underneath
		// it (very likely, since it floats right over the document area)
		// reports the toolbar as the topmost window at that point even
		// though the user's cursor is still logically over their document.
		// Treat that the same as ending back over the start app rather than
		// silently dropping the gesture.
		endKind = startKind
	} else {
		endExe := exePathForWindow(endHwnd)
		endKind = office.DetectAppKind(endExe)
	}
	fmt.Printf("[hook] %s drag detected: start=%q (%q) end kind=%q\n", ts(), startExe, startKind, endKind)
	if startKind == "" || startKind != endKind {
		return
	}
	office.RequestHighlight(startKind, ui.CurrentColor().Hex, office.Point{X: dragStartPt.X, Y: dragStartPt.Y}, office.Point{X: endPt.X, Y: endPt.Y})
}

// Install registers the global low-level mouse hook. It must be called from
// the same OS thread that will keep pumping a Win32 message loop (via
// GetMessage/DispatchMessage) for as long as the hook should stay active —
// low-level hooks are delivered through that thread's message queue.
func Install() {
	cb := syscall.NewCallback(hookProc)
	h, _, errno := procSetWindowsHookExW.Call(whMouseLL, cb, 0, 0)
	hHook = syscall.Handle(h)
	if hHook == 0 {
		fmt.Printf("[hook] SetWindowsHookExW FAILED, no drag will ever be detected: %v\n", errno)
		return
	}
	go processEvents()
	fmt.Println("[hook] global mouse hook installed")
}

// Uninstall removes the hook. Safe to call even if Install failed/wasn't called.
func Uninstall() {
	if hHook != 0 {
		procUnhookWindowsHookEx.Call(uintptr(hHook))
		hHook = 0
	}
}
