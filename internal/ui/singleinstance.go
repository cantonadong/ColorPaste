//go:build windows

package ui

import (
	"syscall"
	"unsafe"
)

// singleInstanceMutexName must stay constant across builds — it's the whole
// point of a named mutex, so every launch of this app (whatever its path)
// checks the same name.
const singleInstanceMutexName = "ColorPaste-SingleInstance-Mutex-9F3B7C1A"

// EnsureSingleInstance returns true if this is the only running instance
// (having claimed the named mutex, held for the rest of the process
// lifetime). If another instance is already running, it brings that
// instance's window to the foreground, shows a message box explaining why,
// and returns false — the caller should exit immediately without creating
// its own window or installing the mouse hook, so two instances never fight
// over the same toolbar position or the same global hook.
func EnsureSingleInstance() bool {
	namePtr := utf16ptr(singleInstanceMutexName)
	// The mutex's own last-error, from this exact call, tells us whether it
	// already existed — a separate syscall.GetLastError() call afterward is
	// not safe to rely on here, since any intervening call (even ones Go's
	// runtime might make) can clobber the thread-local last-error value
	// before we get to read it.
	h, _, errno := procCreateMutexW.Call(0, 0, uintptr(unsafe.Pointer(namePtr)))
	if h == 0 {
		// Couldn't even create the mutex; fail open rather than block the
		// user from ever starting the app.
		return true
	}
	if errno != syscall.Errno(errorAlreadyExists) {
		return true
	}

	if hwnd, _, _ := procFindWindowW.Call(uintptr(unsafe.Pointer(utf16ptr(className))), 0); hwnd != 0 {
		procShowWindow.Call(hwnd, swRestore)
		procSetForegroundWindow.Call(hwnd)
	}

	procMessageBoxW.Call(0,
		uintptr(unsafe.Pointer(utf16ptr("ColorPaste 已经在运行了，正在把它切换到前台。"))),
		uintptr(unsafe.Pointer(utf16ptr("ColorPaste"))),
		mbOK|mbIconInfo|mbTopMost,
	)
	return false
}
