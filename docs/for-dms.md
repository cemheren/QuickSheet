# QuickSheet for dungeon masters

If you're running a tabletop RPG — D&D, Pathfinder, Blades in the Dark, or any
system with dice and initiative — **your desktop wallpaper can be the GM screen.**
No browser tab. No VTT loading screen. Glance down and your encounter state is right
there, behind every window.

> Already familiar with Roll20, Foundry VTT, or Improved Initiative? Think "those,
> but on the actual wallpaper, offline-first, no subscription, and you control
> everything in plain CSV cells."

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/dm-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Your encounter state survives reboots —
the grid is exactly where you left it.

## What goes on the wallpaper

The starter sheet at `examples/dm-dashboard.csv` gives you a GM screen with three
zones you can resize or rearrange:

| Zone              | What it shows                                             | How                                                            |
|-------------------|-----------------------------------------------------------|----------------------------------------------------------------|
| **Initiative**    | Sorted combatant list, current turn marker, round counter | [`init`](https://github.com/cemheren/quicksheet-init-ext)      |
| **Dice**         | Roll results — `2d6+3`, `1d20`, `4d8` — inline            | [`roll`](https://github.com/cemheren/quicksheet-roll-ext)      |
| **Encounter**    | Monster HP, AC, conditions — plain editable cells          | Direct cell editing                                            |
| **Notes**        | Session notes, NPC names, plot hooks                       | Plain cells — type freely                                      |

Both extensions install with a single cell:

```
ext: github:cemheren/quicksheet-roll-ext
ext: github:cemheren/quicksheet-init-ext
```

## Running combat

1. **Set up combatants** — fill the initiative zone with names and initiative rolls.
2. **Roll dice** — type `roll: 1d20+5` in any cell to get an instant result.
3. **Track HP** — decrement monster HP cells manually as damage lands.
4. **Next turn** — the `init:` extension cycles through combatants in order.
5. **Encounter tables** — `roll: table goblin_loot` pulls from a CSV lookup table
   you define (see the roll extension README).

## Why wallpaper > VTT for the GM

- **Always visible** — no alt-tabbing to "check initiative." It's on-screen when
  you glance at your laptop between player turns.
- **Offline** — no server, no internet dependency mid-session.
- **Zero setup per session** — autosaved CSV opens to last state. Start a new
  encounter by editing cells, not navigating UI.
- **Hackable** — extensions are 50-line scripts speaking JSON-lines. Need a custom
  "wild magic surge" table? Write it in Python in 10 minutes.
- **Survives reboots** — your encounter state is a file on disk. Back it up, version
  it, share it.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Nord, Dracula, Matrix, Synthwave. Matrix or
Dracula fit the "fantasy dungeon" vibe. Nord works for a cleaner parchment look.

## Recipe: the minimal GM screen

1. **Column A** — combatant names + initiative order (`init:` prefix).
2. **Column C** — dice roller cell (`roll: 1d20`). Change the formula and press
   Enter to re-roll.
3. **Columns E–H** — monster stat block (Name, HP, AC, Conditions). Edit freely
   during combat.
4. **Row 12+** — session notes. URLs become clickable hyperlinks (monster stat
   blocks on D&D Beyond, maps, etc.).

## Extensions that pair well

- [`roll`](https://github.com/cemheren/quicksheet-roll-ext) — dice notation
  parser + encounter-table CSV lookup.
- [`init`](https://github.com/cemheren/quicksheet-init-ext) — initiative tracker
  with sorted combatant list, turn cycling, and round counter.
- [`define`](https://github.com/cemheren/quicksheet-define-ext) — quick dictionary
  lookup for rules-lawyering a spell description mid-session.

## Share your setup

Post your GM-screen wallpaper on r/unixporn (the rice angle is real), r/DnD,
r/rpg, or r/DMAcademy. Open an issue if you want a new TTRPG-focused extension.
