// Package palette holds the fixed 16-color highlight palette shown in the
// floating toolbar, in the exact order specified by need.md (4x4, left to
// right, top to bottom).
package palette

// Color is one selectable highlight swatch.
type Color struct {
	Name string
	Hex  string
	R, G, B byte
}

// Colors is the ordered 16-color palette. Index 0 is the default selection.
var Colors = []Color{
	{Name: "鲜红", Hex: "#ffaea6"},
	{Name: "橘红", Hex: "#ffb585"},
	{Name: "橙", Hex: "#ffc27c"},
	{Name: "金黄", Hex: "#ffd772"},
	{Name: "柠檬黄", Hex: "#eae381"},
	{Name: "黄绿", Hex: "#cae897"},
	{Name: "嫩绿", Hex: "#adeeae"},
	{Name: "翠绿", Hex: "#99f0ca"},
	{Name: "青绿", Hex: "#93efe2"},
	{Name: "天蓝", Hex: "#90ecfc"},
	{Name: "湛蓝", Hex: "#9ce2ff"},
	{Name: "蓝紫", Hex: "#bbd7ff"},
	{Name: "紫", Hex: "#dccbff"},
	{Name: "紫红", Hex: "#f9bbfc"},
	{Name: "玫红", Hex: "#ffb4e6"},
	{Name: "粉红", Hex: "#ffb2ca"},
}

func init() {
	for i := range Colors {
		c := &Colors[i]
		c.R, c.G, c.B = parseHex(c.Hex)
	}
}

func parseHex(hex string) (r, g, b byte) {
	// hex is always a literal "#rrggbb" constant above.
	v := func(c byte) byte {
		switch {
		case c >= '0' && c <= '9':
			return c - '0'
		case c >= 'a' && c <= 'f':
			return c - 'a' + 10
		case c >= 'A' && c <= 'F':
			return c - 'A' + 10
		}
		return 0
	}
	r = v(hex[1])<<4 | v(hex[2])
	g = v(hex[3])<<4 | v(hex[4])
	b = v(hex[5])<<4 | v(hex[6])
	return
}
