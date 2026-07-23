//go:build windows

package office

import "github.com/go-ole/go-ole/oleutil"

// applyPowerPoint sets the current text selection's highlight color. Font.
// Highlight only exists on PowerPoint 2019/365+; on older versions (and,
// potentially, older WPS Presentation builds) this property lookup fails and
// PutProperty returns an error, which the caller treats as "app not
// compatible, no effect" per spec. Serves both PowerPoint and WPS
// Presentation, depending on progIDs.
func applyPowerPoint(progIDs []string, hex string) error {
	app, cleanup, err := getActiveApp(progIDs)
	if err != nil {
		return err
	}
	defer cleanup()

	activeWindow, err := getIDispatchProp(app, "ActiveWindow")
	if err != nil {
		return err
	}
	defer activeWindow.Release()

	selection, err := getIDispatchProp(activeWindow, "Selection")
	if err != nil {
		return err
	}
	defer selection.Release()

	textRange, err := getIDispatchProp(selection, "TextRange")
	if err != nil {
		return err
	}
	defer textRange.Release()

	font, err := getIDispatchProp(textRange, "Font")
	if err != nil {
		return err
	}
	defer font.Release()

	_, err = oleutil.PutProperty(font, "Highlight", hexToColorRef(hex))
	return err
}
