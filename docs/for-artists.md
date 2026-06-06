# QuickSheet for artists & designers

Your wallpaper is dead real estate. Turn it into an always-visible reference board:
color palettes, client deadlines, project links, measurement conversions — all painted
*behind* your Photoshop / Figma / Blender windows, never stealing focus.

> Not a design tool. A persistent scratchpad that stays out of the way while you work —
> glanceable context for the creative process.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/artist-dashboard.csv
```

Edit cells live, autosaves to CSV every 5 seconds. Add to startup and your reference
board is there every session — behind every window.

## What goes on the wallpaper

Starter sheet at `examples/artist-dashboard.csv`. Columns:

| Column            | What it shows                                                  | How                                                                                       |
|-------------------|----------------------------------------------------------------|-------------------------------------------------------------------------------------------|
| **Palette**       | HEX ↔ RGB ↔ HSL ↔ CMYK conversion, WCAG contrast ratios       | [`color`](https://github.com/Deskworks/quicksheet-color) extension                        |
| **Clients**       | Deadlines, deliverable status, payment tracking                | Plain cells + `c:red:` / `c:green:` color prefixes                                        |
| **References**    | Clickable links to mood boards, Dribbble shots, Figma files    | Paste any URL — auto-detected as a hyperlink                                               |
| **Measurements**  | Canvas sizes, print DPI, cm ↔ inches                           | [`unit`](https://github.com/Deskworks/quicksheet-unit-ext) extension                      |
| **Quick launch**  | Open Photoshop, Blender, project folders in one click          | `r: code ~/Projects/client-a` or `r: gimp`                                                |
| **Definitions**   | Look up art terms, typography jargon                           | [`def`](https://github.com/Deskworks/quicksheet-define-ext) extension                     |

## Recipes

### Color palette board

Keep your project palette visible at all times:

```
A1: c:cyan: Brand palette
B1: color: #1a1a2e
B2: color: #16213e
B3: color: #0f3460
B4: color: #e94560
B5: color: palette #e94560
```

The `color:` extension shows each color's HEX, RGB, HSL, and CMYK breakdown. The
`palette` sub-command generates complementary/analogous colors from a base.

### Client tracker with deadlines

```
A1: c:yellow: CLIENT        B1: c:yellow: DEADLINE    C1: c:yellow: STATUS
A2: Logo redesign – Acme    B2: 2026-06-15            C2: c:green: delivered
A3: Album art – Nova        B3: 2026-06-22            C3: c:red: in progress
A4: UI kit – StartupCo      B4: 2026-07-01            C4: pending
```

Color-code status cells to spot overdue work at a glance.

### Quick-launch creative apps

```
A1: r: gimp
A2: r: blender
A3: r: inkscape
A4: r: xdg-open ~/References/moodboard.pdf
A5: https://www.figma.com/file/your-project
```

One click to open any tool or reference. URLs become clickable hyperlinks automatically.

### Unit conversion for print work

```
A1: unit: 8.5in to cm
A2: unit: 300dpi * 11in to px
A3: unit: 210mm to in
```

Convert between imperial and metric for print layouts without switching apps.

## What the wallpaper IS NOT for

- **Pixel-perfect design work.** Use Figma / Illustrator / Affinity for that.
- **Asset management.** Use your DAM or file system. The wallpaper is a quick-reference
  layer, not a gallery.
- **Color picking from screen.** Use your OS color picker. The `color:` extension
  converts values you already have.

## Who this is for

- **Freelance illustrators** tracking multiple client projects and deadlines
- **Graphic designers** who need palettes and brand specs always visible
- **3D artists** juggling render settings, reference links, and project timelines
- **UI/UX designers** keeping design tokens and component specs at a glance
- **Photographers** tracking shoot schedules, client deliveries, and print specs

## See also

- [Main README](../README.md) — full feature list and extension catalog
- [For homelabbers](for-homelab.md) — server monitoring dashboard
- [For traders](for-traders.md) — financial watchlist and P/L tracking
- [For students](for-students.md) — coursework and study tools
