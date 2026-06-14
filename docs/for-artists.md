# QuickSheet for artists & creatives

Freelance illustrators, designers, photographers, and musicians juggle clients,
deadlines, palettes, and invoices — often across multiple tools and browser tabs.
**QuickSheet puts your creative admin layer on the wallpaper** so the info you
glance at between brushstrokes is always there, never buried under Photoshop or
your DAW.

> This isn't a project manager. It's the ambient info panel that answers
> "who owes me money, what's due next, what hex is that again" without
> switching windows.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/artist-portfolio.csv
```

Edit cells live, autosaves to CSV every 5 seconds. Add to startup and your
creative dashboard is painted behind every window — never stealing focus from
your canvas.

## What goes on the wallpaper

Starter sheet at `examples/artist-portfolio.csv`. Columns:

| Column            | What it shows                                      | Extension                                                                                   |
|-------------------|----------------------------------------------------|---------------------------------------------------------------------------------------------|
| **Commissions**   | Client name, piece title, status (sketch/WIP/done) | plain cells — just type                                                                     |
| **Deadlines**     | Days until delivery, countdown progress            | [`cntdn`](https://github.com/Deskworks/quicksheet-cntdn) (countdown to any date)            |
| **Palette**       | Hex codes ↔ RGB ↔ HSL for your current project     | [`color`](https://github.com/Deskworks/quicksheet-color) (color converter + palette gen)    |
| **Hourly rate**   | Your minimum viable hourly rate from annual target | [`rate`](https://github.com/Deskworks/quicksheet-rate) (freelance rate calculator)           |
| **Budget**        | Monthly spending envelopes with progress bars      | [`budget`](https://github.com/Deskworks/quicksheet-budget) (envelope visualizer)            |
| **Client zones**  | Current time in each client's timezone             | [`worldtm`](https://github.com/Deskworks/quicksheet-worldtm) (multi-timezone clock)        |
| **Invoices**      | Quarterly estimated tax countdown                  | [`qtr`](https://github.com/Deskworks/quicksheet-qtr) (IRS deadline tracker)                 |

All extensions install with a single `ext: github:Deskworks/quicksheet-<name>` cell.

## Example workflows

### Commission tracker

```
| Client       | Piece            | Status | Due        | Days left        |
|--------------|------------------|--------|------------|------------------|
| @alice_art   | Character design | WIP    | 2025-07-01 | ext: cntdn: Jul 1 |
| @bob_studio  | Album cover      | Sketch | 2025-07-15 | ext: cntdn: Jul 15|
```

### Palette reference row

Keep your project palette visible at all times:

```
ext: color: #2E1065    →  rgb(46,16,101) / hsl(267,73%,23%)
ext: color: #7C3AED    →  rgb(124,58,237) / hsl(263,76%,58%)
ext: color: palette: #2E1065  →  complementary, triadic, analogous
```

### Know-your-rate check

```
ext: rate: 45000,1800   → $25.00/hr minimum (target income / billable hours)
```

If your minimum viable rate is higher than what a client offers, the number
speaks for itself — no awkward negotiation math in your head.

## What this IS NOT

- **Not a portfolio website.** Your Behance/ArtStation/personal site handles that.
- **Not a project manager.** Notion, Trello, and Asana are fine for deep task
  trees. This is the *glanceable summary* — a heads-up display, not a cockpit.
- **Not an invoicing tool.** Use Wave, Stripe, or your accountant for real
  invoices. QuickSheet tells you *when* payments are due, not generates them.

## Getting started

1. Launch with the example CSV above.
2. Replace the sample data with your real commissions and deadlines.
3. Install `color` and `cntdn` extensions (one cell each, they self-install).
4. Optionally add `rate` and `budget` if you freelance full-time.
5. Enjoy glanceable creative admin behind your canvas app of choice.

---

Back to [main README](../README.md) · More guides: [for homelabbers](for-homelab.md) · [for traders](for-traders.md) · [for SREs](for-sre.md) · [for students](for-students.md)
