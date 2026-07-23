//go:build windows

// Package office applies a highlight color to the current text selection in
// an already-running Microsoft Word / Excel / PowerPoint instance via COM
// automation. It never launches these apps itself — if the target app isn't
// already running (or its object model doesn't expose the property we need),
// the request is simply dropped, matching the "unsupported app -> no effect"
// requirement.
package office

import (
	"fmt"
	"path/filepath"
	"runtime"
	"strings"
	"time"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type AppKind string

const (
	Word            AppKind = "word"
	Excel           AppKind = "excel"
	PowerPoint      AppKind = "powerpoint"
	WPSWriter       AppKind = "wps-writer"
	WPSSheet        AppKind = "wps-sheet"
	WPSPresentation AppKind = "wps-presentation"
)

// DetectAppKind maps a process executable name (e.g. from
// QueryFullProcessImageNameW) to the Office app it belongs to, or "" if it's
// not one of our supported apps.
func DetectAppKind(exePath string) AppKind {
	switch strings.ToLower(filepath.Base(exePath)) {
	case "winword.exe":
		return Word
	case "excel.exe":
		return Excel
	case "powerpnt.exe":
		return PowerPoint
	case "wps.exe":
		return WPSWriter
	case "et.exe":
		return WPSSheet
	case "wpp.exe":
		return WPSPresentation
	}
	return ""
}

// progIDsFor lists the COM ProgIDs to try, in order, for a given AppKind.
// WPS's own ProgID naming has varied across releases/builds (older WPS also
// sometimes registers the plain "kwps.Application" family under different
// casings), so each WPS kind tries a short list of known candidates rather
// than a single hardcoded name.
func progIDsFor(kind AppKind) []string {
	switch kind {
	case Word:
		return []string{"Word.Application"}
	case Excel:
		return []string{"Excel.Application"}
	case PowerPoint:
		return []string{"PowerPoint.Application"}
	case WPSWriter:
		return []string{"kwps.Application", "KWPS.Application", "wps.Application"}
	case WPSSheet:
		return []string{"ket.Application", "KET.Application", "et.Application"}
	case WPSPresentation:
		return []string{"kwpp.Application", "KWPP.Application", "wpp.Application"}
	}
	return nil
}

// Point is a screen coordinate in physical pixels, as reported by the global
// mouse hook (and, since this process is now DPI-aware, consistent with
// every other coordinate-based Win32/COM call this package and internal/ui
// make).
type Point struct{ X, Y int32 }

type request struct {
	kind       AppKind
	hex        string
	start, end Point
}

// requests holds at most one pending request: apply() always reads whatever
// is *currently* selected at call time rather than anything captured at
// drag-end, so a backlog is actively harmful, not just wasted work — a
// request that sits queued behind a slow COM call ends up shading whatever
// the user has selected by the time its turn comes, which after a couple of
// quick successive drags is no longer the selection that request was for.
// RequestHighlight keeps only the newest request pending so the worker is
// never more than one drag behind.
var requests = make(chan request, 1)

func init() {
	go worker()
}

// worker owns every COM call on one dedicated, apartment-initialized OS
// thread — Office's COM objects are STA-threaded, so all calls into them
// must consistently come from the same thread.
func worker() {
	runtime.LockOSThread()
	if err := ole.CoInitialize(0); err != nil {
		fmt.Println("[office] CoInitialize failed:", err)
		return
	}
	defer ole.CoUninitialize()

	for req := range requests {
		t0 := time.Now()
		fmt.Printf("[office] %s worker picked up %s request\n", ts(), req.kind)
		if err := apply(req.kind, req.hex, req.start, req.end); err != nil {
			fmt.Printf("[office] %s highlight not applied (app not open, or incompatible) after %s: %v\n", ts(), since(t0), err)
		} else {
			fmt.Printf("[office] %s applied %s highlight to current %s selection (took %s)\n", ts(), req.hex, req.kind, since(t0))
		}
	}
}

// ts returns a millisecond-precision timestamp for correlating log lines
// against when the user actually did something.
func ts() string { return time.Now().Format("15:04:05.000") }

func since(t0 time.Time) string { return time.Since(t0).Round(time.Millisecond).String() }

// RequestHighlight asynchronously asks the COM worker to apply hex (e.g.
// "#ffaea6") to the drag from start to end (screen coordinates, physical
// pixels) in the given already-running app. Safe to call from any goroutine
// — including a low-level mouse hook callback — since it never blocks
// waiting on COM.
func RequestHighlight(kind AppKind, hex string, start, end Point) {
	for {
		select {
		case requests <- request{kind, hex, start, end}:
			return
		default:
			// A previous request is still sitting there unprocessed — it's
			// now stale (see the comment on requests), so drop it to make
			// room instead of leaving it to be applied late and wrong.
			select {
			case <-requests:
			default:
			}
		}
	}
}

func apply(kind AppKind, hex string, start, end Point) error {
	progIDs := progIDsFor(kind)
	if progIDs == nil {
		return fmt.Errorf("unsupported app kind %q", kind)
	}
	switch kind {
	case Word, WPSWriter:
		return applyWord(progIDs, hex, start, end)
	case Excel, WPSSheet:
		return applyExcel(progIDs, hex, start, end)
	case PowerPoint, WPSPresentation:
		return applyPowerPoint(progIDs, hex)
	}
	return fmt.Errorf("unsupported app kind %q", kind)
}

// hexToColorRef converts "#rrggbb" to the BGR-packed COLORREF/OLE_COLOR
// integer Office's object model expects (0x00BBGGRR, same packing Windows
// itself uses for RGB()).
func hexToColorRef(hex string) int32 {
	v := func(c byte) int32 {
		switch {
		case c >= '0' && c <= '9':
			return int32(c - '0')
		case c >= 'a' && c <= 'f':
			return int32(c-'a') + 10
		case c >= 'A' && c <= 'F':
			return int32(c-'A') + 10
		}
		return 0
	}
	r := v(hex[1])<<4 | v(hex[2])
	g := v(hex[3])<<4 | v(hex[4])
	b := v(hex[5])<<4 | v(hex[6])
	return r | g<<8 | b<<16
}

// getActiveApp attaches to an already-running instance registered in the
// Running Object Table, trying each progID in turn (needed for WPS, whose
// registered ProgID name has varied across releases). It deliberately never
// creates a new instance — we only want to act on an app the user already
// has open.
func getActiveApp(progIDs []string) (*ole.IDispatch, func(), error) {
	var lastErr error
	for _, progID := range progIDs {
		t0 := time.Now()
		unknown, err := oleutil.GetActiveObject(progID)
		if d := time.Since(t0); d > 10*time.Millisecond {
			fmt.Printf("[office] %s GetActiveObject(%s) took %s\n", ts(), progID, d.Round(time.Millisecond))
		}
		if err != nil {
			lastErr = err
			continue
		}
		disp, err := unknown.QueryInterface(ole.IID_IDispatch)
		if err != nil {
			unknown.Release()
			lastErr = err
			continue
		}
		return disp, func() {
			disp.Release()
			unknown.Release()
		}, nil
	}
	return nil, func() {}, fmt.Errorf("no active object among %v: %w", progIDs, lastErr)
}

func getIDispatchProp(disp *ole.IDispatch, name string) (*ole.IDispatch, error) {
	v, err := oleutil.GetProperty(disp, name)
	if err != nil {
		return nil, fmt.Errorf("%s: %w", name, err)
	}
	d := v.ToIDispatch()
	if d == nil {
		return nil, fmt.Errorf("%s: not an object", name)
	}
	return d, nil
}
