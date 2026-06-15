# QuickSheet for dungeon masters

Running a TTRPG session? **Your desktop wallpaper becomes the DM screen.** Player stats,
initiative order, dice rolls, encounter tables, session notes — all visible at a glance
behind your VTT or chat window.

> Already using pen-and-paper, Notion, or a GM binder? Think "those, but rendered on the
> actual wallpaper so you never lose focus switching tabs mid-combat."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/dm-screen.csv
```

Edit cells live during the session. Autosaves every 5 seconds — no "oops I forgot to save
the initiative tracker" moments. Add to startup and you have an always-ready DM screen.

## What goes on the wallpaper

The starter sheet at `examples/dm-screen.csv` is a ready-to-use DM screen. The sections:

| Section           | What it shows                                          | How it works                                |
|-------------------|--------------------------------------------------------|---------------------------------------------|
| **Initiative**    | Turn order, HP, AC, status conditions                  | Plain cells — edit live during combat        |
| **Dice**          | Roll any dice inline: d20, 2d6+3, 4d8                  | [`roll`](https://github.com/cemheren/quicksheet-roll-ext) extension |
| **Encounter**     | Random encounter table with weighted rolls             | `roll` extension table lookup               |
| **Party stats**   | Player names, HP, AC, passive perception               | Plain cells, column sums for party totals   |
| **Session notes** | Freeform notes, NPC names, plot threads                | Plain cells — your scratch pad              |
| **Quick refs**    | Conditions, DC guidelines, cover rules                 | Plain cells with rule summaries             |

## Install the roll extension

One cell installs it:

```
ext: github:cemheren/quicksheet-roll-ext
```

Then use the `roll:` prefix in any cell:

```
roll: d20          → 14
roll: 2d6+3       → 11
roll: 4d8          → 19
roll: d100         → 73
```

Press **Enter** on a roll cell to re-roll. The result updates in place.

## Make it yours

Cycle themes with **Ctrl+T** — Matrix or Dracula fit the dark-tavern vibe. Gruvbox
for a parchment look.

Layout tips:
- **Left columns**: initiative tracker (changes every round)
- **Center**: dice + encounter tables (reference during play)
- **Right columns**: session notes + party stats (rarely changes mid-combat)

## Why wallpaper > a second monitor or tab

- **Always visible** — behind your VTT, chat, or music player. Glance down, see HP.
- **No alt-tab** — initiative is visible even while sharing your screen for maps.
- **Survives reboots** — your DM prep is there on boot. No "which Notion page was it?"
- **Offline** — no internet needed. Your CSV is a local file.
- **Editable mid-session** — click a cell, type new HP, move on. No app chrome.

## Recipe: initiative tracker

1. Column A: turn order (1, 2, 3…)
2. Column B: character/monster name
3. Column C: HP (edit as damage happens)
4. Column D: AC (reference)
5. Column E: status conditions (Prone, Concentrating, etc.)
6. Sort manually by initiative roll at combat start. Rearrange rows as needed.

## Recipe: random encounter table

Use `roll:` with a lookup pattern:

```
roll: d6    → 3
```

Then reference a table in adjacent cells:

| d6 | Encounter          |
|----|-------------------|
| 1  | 2d4 goblins       |
| 2  | 1 owlbear         |
| 3  | Merchant caravan   |
| 4  | Bandit ambush      |
| 5  | Nothing            |
| 6  | Young dragon       |

## More extensions to mix in

- [`cal`](https://github.com/cemheren/quicksheet-cal-ext) — in-world calendar tracking
- [`epoch`](https://github.com/Deskworks/quicksheet-epoch-ext) — session timer (how long have we been playing?)
- [`words`](https://github.com/cemheren/quicksheet-words) — word count for session recap notes

See the [full extension directory](extensions.md).

## Share your setup

Post your DM-screen wallpaper on r/DnD, r/DMAcademy, r/unixporn (the rice is real),
or r/rpg. Open an issue if you want a new TTRPG-specific extension listed.
