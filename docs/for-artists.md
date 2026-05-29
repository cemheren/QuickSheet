# QuickSheet for artists & creatives

If you draw, paint, animate, write, make music, or build worlds — **your desktop
wallpaper can be your project dashboard.** Commission tracker, deadline countdown,
colour palette reference, word-count progress, inspiration prompts — all painted
*behind* your creative tools, always visible, never stealing focus.

> Already using sticky notes, Notion sidebars, or second-monitor spreadsheets to
> track projects? Think "that, but embedded in the wallpaper itself, updating live,
> and stored as a plain CSV you own."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/artist-dashboard.csv
```

Edit cells, autosaves to CSV every 5 seconds. Add to startup and the grid is there
every session — *behind* Photoshop, Clip Studio, your DAW, your IDE.

## What goes on the wallpaper

Starter sheet at `examples/artist-dashboard.csv`. Sections:

| Section             | What it shows                                               | How                                                                                       |
|---------------------|-------------------------------------------------------------|-------------------------------------------------------------------------------------------|
| **Commissions**     | Client, description, status, deadline, price                | Plain cells — edit in place, colour-code by status                                        |
| **Word count**      | Live WC + reading time for your draft file                  | [`words`](https://github.com/cemheren/quicksheet-words) extension                         |
| **Deadlines**       | Days remaining until each due date                          | Inline formula cells or date arithmetic                                                   |
| **Palette**         | Hex codes + colour names for your current project palette   | Plain cells — visual reference strip                                                      |
| **Inspiration**     | Random prompt or quote refreshed on demand                  | `i: shuf -n1 ~/prompts.txt` (inline process)                                             |
| **AI brainstorm**   | Quick concept generation, description refinement            | [`ollama`](https://github.com/cemheren/quicksheet-ollama) extension (local, private)      |

## Use cases by discipline

### Illustrators & digital painters

- **Commission queue**: one row per commission — client name, brief, deadline, payment
  status. Sort by deadline. Mark cells red when overdue.
- **Colour palette strip**: a row of hex codes (`#2E4057`, `#048A81`, ...) as quick
  reference while painting. Swap the row per project.
- **Reference links**: paste ArtStation / Pinterest URLs — QuickSheet auto-detects
  them as clickable hyperlinks. One click opens the ref.

### Writers & worldbuilders

- **Word count tracker**: point the `words:` extension at your manuscript file.
  See live word count, reading time, and Flesch readability without leaving your editor.
- **Chapter progress**: one row per chapter, track draft/edit/done status.
- **Name generator**: `i: shuf -n3 ~/names.txt | paste -sd ', '` gives you three
  random character names on refresh.

### Musicians & composers

- **Session log**: date, BPM, key, notes for each recording session.
- **Set list**: track order, duration, transitions — plain CSV, easy to reorder.
- **Tuning reference**: standard tunings, capo charts, or scale intervals as a
  permanent wallpaper reference strip.

### Animators & game devs

- **Shot list / asset tracker**: scene, frame count, status, assignee.
- **Build status**: pair with [`gha`](https://github.com/cemheren/quicksheet-gha-ext)
  to see CI status for your game project.
- **Version notes**: keep a running changelog visible while you work.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand. Synthwave and Dracula fit creative
aesthetics particularly well.

## Why wallpaper > a second app

- **Always visible** — no switching to Notion or Trello mid-flow. Your project
  state is literally the background.
- **Zero distraction** — no notifications, no feeds, no social. Just your data.
- **Survives reboots** — CSV autosave, same grid every boot.
- **No account, no SaaS** — your commission list stays on your machine.
- **Infinitely customisable** — plain text cells, inline commands, extensions.
  If you can script it, it can live on your wallpaper.

## Recipe: commission tracker

1. Row per commission: `Client | Brief | Deadline | Price | Status`
2. Status column: type `done`, `wip`, `waiting`, `overdue` as plain text.
3. Add a header row with column names.
4. Sort manually or use search (Ctrl+F) to find clients.
5. When a commission is done, move the row down or mark it — your CSV, your rules.

## Extensions that pair well

- [`words`](https://github.com/cemheren/quicksheet-words) — word/character count
  and readability for writers.
- [`ollama`](https://github.com/cemheren/quicksheet-ollama) — local AI for
  brainstorming, descriptions, naming. Private, no API key needed.
- [`ghstreak`](https://github.com/cemheren/quicksheet-ghstreak-ext) — if you
  publish creative code (shaders, generative art, plugins).
- [`unitconv`](https://github.com/cemheren/quicksheet-unitconv) — convert between
  px/in/cm/mm for print sizing.

See the [full extension directory](extensions.md).

## Share your setup

Post your creative wallpaper on r/unixporn, r/DigitalArt, or r/worldbuilding.
Open an issue if you want a new prefix or extension officially listed.
