//go:build windows

package office

import (
	"fmt"

	"github.com/go-ole/go-ole/oleutil"
)

// applyExcel fills the currently selected cells' interior with hex — Excel
// has no separate "text highlight", so cell background is the closest
// equivalent to what Word/PowerPoint call highlighting. Serves both Excel
// and WPS Sheet, depending on progIDs.
//
// Unlike Word, this reads Selection rather than hit-testing start/end via
// Window.RangeFromPoint — Excel.Window.RangeFromPoint has a known Microsoft
// bug where it returns Nothing on displays scaled above 100% (confirmed:
// this machine runs at 150%), regardless of how correct the screen
// coordinates passed to it are. start/end are accepted for symmetry with
// applyWord/the shared apply() dispatch but unused here.
func applyExcel(progIDs []string, hex string, _, _ Point) error {
	app, cleanup, err := getActiveApp(progIDs)
	if err != nil {
		return err
	}
	defer cleanup()

	selection, err := getIDispatchProp(app, "Selection")
	if err != nil {
		return err
	}
	defer selection.Release()

	if addrVar, aerr := oleutil.GetProperty(selection, "Address"); aerr == nil {
		fmt.Printf("[office] %s excel selection=%s\n", ts(), addrVar.ToString())
	}

	interior, err := getIDispatchProp(selection, "Interior")
	if err != nil {
		return err
	}
	defer interior.Release()

	_, err = oleutil.PutProperty(interior, "Color", hexToColorRef(hex))
	return err
}
