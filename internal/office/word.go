//go:build windows

package office

import (
	"fmt"

	ole "github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// applyWord shades hex onto the text between start and end (screen
// coordinates, physical pixels, as captured by the drag gesture that
// triggered this request), in whichever of progIDs is the currently running
// instance — this same function serves both Microsoft Word (progIDs =
// ["Word.Application"]) and WPS Writer (progIDs = WPS's candidate ProgIDs),
// since WPS's object model mirrors Word's Selection/Range/Shading/Paragraphs
// closely enough for macro compatibility. Word's classic "text highlight"
// (HighlightColorIndex) only supports 15 fixed system colors, so we use
// character Shading instead — the same mechanism modern Word itself falls
// back to for custom (non-standard) highlight colors.
func applyWord(progIDs []string, hex string, start, end Point) error {
	app, cleanup, err := getActiveApp(progIDs)
	if err != nil {
		return err
	}
	defer cleanup()

	win, err := getIDispatchProp(app, "ActiveWindow")
	if err != nil {
		return err
	}
	defer win.Release()

	// RangeFromPoint hit-tests screen coordinates directly into document
	// character positions, deliberately never touching Selection. Selection
	// is what Word's own "smart" drag-selection heuristics operate on (e.g.
	// snapping to whole paragraphs once a drag crosses a paragraph
	// boundary) — turning that off in Word's own editing options didn't
	// stop it happening, so the only reliable way to shade exactly what was
	// dragged, no more and no less, is to derive the range straight from the
	// two endpoints ourselves.
	startRng, err := rangeFromPoint(win, start.X, start.Y)
	if err != nil {
		return err
	}
	defer startRng.Release()

	endRng, err := rangeFromPoint(win, end.X, end.Y)
	if err != nil {
		return err
	}
	defer endRng.Release()

	s1, err := intProp(startRng, "Start")
	if err != nil {
		return err
	}
	e1, err := intProp(startRng, "End")
	if err != nil {
		return err
	}
	s2, err := intProp(endRng, "Start")
	if err != nil {
		return err
	}
	e2, err := intProp(endRng, "End")
	if err != nil {
		return err
	}

	// The drag can run in either direction (left-to-right, right-to-left,
	// bottom-to-top...), so the endpoint reached last isn't necessarily the
	// later position in the document — take the outer bounds instead.
	lo, hi := s1, e1
	if s2 < lo {
		lo = s2
	}
	if e2 > hi {
		hi = e2
	}
	fmt.Printf("[office] %s RangeFromPoint start=(%d,%d)->[%d,%d) end=(%d,%d)->[%d,%d) combined=[%d,%d)\n",
		ts(), start.X, start.Y, s1, e1, end.X, end.Y, s2, e2, lo, hi)

	doc, err := getIDispatchProp(app, "ActiveDocument")
	if err != nil {
		return err
	}
	defer doc.Release()

	rngVar, err := oleutil.CallMethod(doc, "Range", lo, hi)
	if err != nil {
		return fmt.Errorf("Document.Range: %w", err)
	}
	rng := rngVar.ToIDispatch()
	if rng == nil {
		return fmt.Errorf("Document.Range: not an object")
	}
	defer rng.Release()

	if textVar, terr := oleutil.GetProperty(rng, "Text"); terr == nil {
		text := textVar.ToString()
		if len(text) > 60 {
			text = text[:60] + "..."
		}
		fmt.Printf("[office] %s word range [%d,%d) text=%q len=%d\n", ts(), lo, hi, text, len(textVar.ToString()))
	}

	return shadeExcludingParagraphMarks(doc, rng, lo, hi, hex)
}

// shadeExcludingParagraphMarks applies hex to [lo,hi) one paragraph at a
// time, each capped just before that paragraph's own trailing mark.
// Confirmed by hand (dragging a small selection that crosses a paragraph
// boundary, no scrolling involved, both paragraphs already fully visible):
// shading a single Range that happens to include a paragraph mark makes Word
// render the *whole* paragraph as shaded, not just the characters actually
// in the range — Word's own "Text Highlight Color" tool doesn't have this
// problem on the exact same drag, which means it isn't reading the mark's
// formatting as representative of the whole paragraph the way straight
// Range.Shading does. Never letting any single Shading call touch a
// paragraph mark avoids it entirely.
func shadeExcludingParagraphMarks(doc, rng *ole.IDispatch, lo, hi int32, hex string) error {
	paragraphs, err := getIDispatchProp(rng, "Paragraphs")
	if err != nil {
		return err
	}
	defer paragraphs.Release()

	count, err := intProp(paragraphs, "Count")
	if err != nil {
		return err
	}
	fmt.Printf("[office] %s shading [%d,%d) across %d paragraph(s)\n", ts(), lo, hi, count)

	colorRef := hexToColorRef(hex)
	for i := 1; i <= int(count); i++ {
		paraVar, err := oleutil.CallMethod(paragraphs, "Item", i)
		if err != nil {
			return fmt.Errorf("Paragraphs.Item(%d): %w", i, err)
		}
		para := paraVar.ToIDispatch()
		if para == nil {
			return fmt.Errorf("Paragraphs.Item(%d): not an object", i)
		}

		paraRange, err := getIDispatchProp(para, "Range")
		if err != nil {
			para.Release()
			return err
		}
		pStart, e1 := intProp(paraRange, "Start")
		pEnd, e2 := intProp(paraRange, "End")
		paraRange.Release()
		para.Release()
		if e1 != nil {
			return e1
		}
		if e2 != nil {
			return e2
		}

		subStart, subEnd := lo, hi
		if pStart > subStart {
			subStart = pStart
		}
		if pEnd-1 < subEnd {
			subEnd = pEnd - 1 // exclude this paragraph's own trailing mark
		}
		fmt.Printf("[office] %s   para %d: range=[%d,%d) sub=[%d,%d)\n", ts(), i, pStart, pEnd, subStart, subEnd)
		if subEnd <= subStart {
			continue // nothing but the mark itself in this paragraph
		}

		subRngVar, err := oleutil.CallMethod(doc, "Range", subStart, subEnd)
		if err != nil {
			return fmt.Errorf("Document.Range(%d,%d): %w", subStart, subEnd, err)
		}
		subRng := subRngVar.ToIDispatch()
		if subRng == nil {
			return fmt.Errorf("Document.Range(%d,%d): not an object", subStart, subEnd)
		}

		shading, err := getIDispatchProp(subRng, "Shading")
		if err != nil {
			subRng.Release()
			return err
		}
		_, err = oleutil.PutProperty(shading, "BackgroundPatternColor", colorRef)
		shading.Release()
		subRng.Release()
		if err != nil {
			return err
		}
	}
	return nil
}

func rangeFromPoint(win *ole.IDispatch, x, y int32) (*ole.IDispatch, error) {
	v, err := oleutil.CallMethod(win, "RangeFromPoint", x, y)
	if err != nil {
		return nil, fmt.Errorf("RangeFromPoint: %w", err)
	}
	d := v.ToIDispatch()
	if d == nil {
		return nil, fmt.Errorf("RangeFromPoint: not an object")
	}
	return d, nil
}

func intProp(disp *ole.IDispatch, name string) (int32, error) {
	v, err := oleutil.GetProperty(disp, name)
	if err != nil {
		return 0, fmt.Errorf("%s: %w", name, err)
	}
	switch n := v.Value().(type) {
	case int32:
		return n, nil
	case int64:
		return int32(n), nil
	case int:
		return int32(n), nil
	default:
		return 0, fmt.Errorf("%s: unexpected type %T", name, v.Value())
	}
}
