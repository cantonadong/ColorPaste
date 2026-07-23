// Command colorpaste is the ColorPaste floating color toolbar.
//
// See docs/需求开发文档.md. Current status: toolbar UI (steps 1-3) plus a
// first working slice of Office highlight automation (steps 4-5) — dragging
// over text in an already-open Word/Excel/PowerPoint window while a color is
// selected shades that selection with the chosen color.
package main

import (
	"fmt"
	"os"

	"colorpaste/internal/hook"
	"colorpaste/internal/palette"
	"colorpaste/internal/ui"
)

func main() {
	fmt.Println("ColorPaste — click a swatch to select, then drag over text in Word/Excel/PowerPoint to highlight it.")

	if !ui.EnsureSingleInstance() {
		return
	}

	err := ui.RunToolbarPreview(
		func(c palette.Color) {},
		func() {
			hook.Install()
		},
	)
	hook.Uninstall()
	if err != nil {
		fmt.Fprintln(os.Stderr, "error:", err)
		os.Exit(1)
	}
}
