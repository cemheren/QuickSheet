# QuickSheet for Dungeon Masters

If you're running a tabletop campaign — D&D, Pathfinder, Fate, or any system — **your
desktop wallpaper can be the DM screen.** Initiative tracker, dice roller, HP bars, loot
tables, and session notes — always visible behind Roll20 or your VTT, no extra tabs.

> Already familiar with D&D Beyond, Notion DM templates, or OneNote? Think "those, but on
> the actual wallpaper, offline, and talking to dice-roller extensions you can write in
> 50 lines of any language."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/dm-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Survive reboots — your campaign state is
exactly where you left it.

## What goes on the wallpaper

The starter sheet at `examples/dm-dashboard.csv` gives you a ready-to-use DM screen:

| Zone              | What it shows                                  | How                                                              |
|-------------------|------------------------------------------------|------------------------------------------------------------------|
| **Party tracker** | Name, HP, AC, conditions                       | Plain cells — edit mid-combat                                    |
| **Dice roller**   | Any dice expression, live results              | [`roll`](https://github.com/cemheren/quicksheet-roll-ext)        |
| **Initiative**    | Sorted turn order for encounters               | [`init`](https://github.com/cemheren/quicksheet-init-ext)        |
| **Loot table**    | Items, value, who claimed what                 | Plain cells                                                      |
| **Session notes** | Running log of what happened                   | Plain cells                                                      |

Extensions install with a single `ext: github:cemheren/quicksheet-roll-ext` cell —
QuickSheet clones the repo and the prefix is live instantly.

## Dice rolling on the wallpaper

The [`roll`](https://github.com/cemheren/quicksheet-roll-ext) extension handles standard
dice notation:

```
ext: github:cemheren/quicksheet-roll-ext
---
roll: 1d20+5
roll: 2d6+3
roll: 4d6kh3
roll: 1d100
roll: 8d6
```

Every cell shows a live result. Press Enter on a roll cell to re-roll. `4d6kh3` ("keep
highest 3") is the classic stat-generation formula — useful for session-zero character
creation.

## Initiative tracking

The [`init`](https://github.com/cemheren/quicksheet-init-ext) extension sorts combatants
by roll:

```
ext: github:cemheren/quicksheet-init-ext
---
init: Kira,18
init: Bron,12
init: Stone Golem,9
init: Shadow Wraith,15
init: Nyx,20
```

Sorted output appears in order. Update a value, re-roll initiative, the list re-sorts.

## Combat tracker layout

A practical mid-combat layout:

```
Party,HP,AC,Status
Kira (Rogue),38/45,16,—
Bron (Paladin),52/67,20,Blessed
Zeph (Wizard),22/30,13,Concentrating
Nyx (Ranger),41/41,15,Hunter's Mark
```

Edit HP in place as damage lands. Ctrl+Z if you fat-finger it. The grid is your scratch
paper — no forms, no dropdowns, just cells.

## Session notes that survive crashes

Every 5 seconds the CSV autosaves. Your session notes are never lost — no "forgot to
save the OneNote page" moments. Keep a running log:

```
Session 12 Notes
Party entered vault via trapped corridor
Kira disarmed 2/3 traps
Bron triggered trap #3 (8 dmg)
Next: deeper vault + boss fight
```

After the session, the CSV is a plain-text artifact you can grep, diff, or commit to a
campaign git repo.

## Make it yours

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula, Synthwave,
Gruvbox, Monokai. Matrix is atmospheric for underdark campaigns; Nord for frost/winter arcs.

## Why wallpaper > app

- **Always visible** — initiative and HP are behind your VTT window, never buried in a tab.
- **Offline** — no internet needed. Works at the game store, on the bus, anywhere.
- **Plain CSV** — your campaign data is a text file. Version it with git, diff between
  sessions, grep for that NPC name you forgot.
- **Hackable** — need a custom loot-table roller for your homebrew system? Write a 50-line
  extension ([extension-protocol.md](extension-protocol.md)).
- **Zero subscription** — no D&D Beyond tier, no Notion plan. Just a CSV on your drive.

## Recipe: full DM screen for a combat encounter

1. Top rows: party tracker (name, HP, AC, conditions).
2. Middle rows: enemy stat block (name, HP, AC, initiative roll cell).
3. Side column: quick-roll cells (`roll: 1d20`, `roll: 2d6`, `roll: 1d8`).
4. Bottom rows: session notes — append as the session progresses.
5. Add a `L: A1, 0` loop cell to keep dice rolls fresh on demand.

## See also

- [Full extension directory](extensions.md)
- [Keyboard shortcuts](keyboard-shortcuts.md)
- [Extension protocol spec](extension-protocol.md) — write your own in an afternoon
- [for-students.md](for-students.md) — if you're also in school
- [for-homelab.md](for-homelab.md) — if you're also running game servers
