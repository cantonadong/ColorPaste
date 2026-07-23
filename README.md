# ColorPaste

<img width="1353" height="643" alt="image" src="https://github.com/user-attachments/assets/27e3e6b3-b2d1-429b-a844-9c5137bf54e5" />


A lightweight Windows desktop tool that lets you highlight text in Word, Excel,
and PowerPoint with any of 16 colors — without leaving your keyboard workflow.

Pick a color from the floating toolbar, then drag across text in an already-open
Office window to apply that highlight color, just as if you'd selected the text
and clicked the app's own highlight button.

## How it works

1. Launch ColorPaste — a small always-on-top 4x4 color palette appears.
2. Click a swatch to select the current highlight color.
3. Drag across text in Word / Excel / PowerPoint (outside the toolbar) to apply
   that color as a highlight/shading.
4. If the target app isn't supported or doesn't expose a highlight API, the drag
   has no effect — it never interferes with normal selection/drag behavior.

## Requirements

- Windows 10+
- Go 1.26+ (to build from source)
- Microsoft Word / Excel / PowerPoint already running (for highlight automation)

## Build

```
go build -ldflags="-H=windowsgui" -o bin/ColorPaste.exe ./cmd/colorpaste
```

The `-H=windowsgui` flag builds it as a GUI subsystem binary so no console/log
window appears on launch.

Copy `assets/icon.ico` next to the built exe (as `assets/icon.ico`) so the app
can load its window/tray icon.

## Project layout

- `cmd/colorpaste` — entry point
- `internal/ui` — Win32 floating toolbar UI
- `internal/palette` — the 16-color palette
- `internal/hook` — global low-level mouse hook for drag detection
- `internal/office` — Word/Excel/PowerPoint COM automation for applying highlights

## Status

Early-stage; see `log.md` for detailed development notes.
