# QuickSheet for Dungeon Masters & GMs

Running a TTRPG session? **Your desktop wallpaper becomes your GM screen.** Initiative
order, dice rolls, NPC hit points, loot tables — all visible at a glance behind your
notes app, no tab-switching, no alt-tabbing away from your VTT.

> Already using a VTT like Foundry, Roll20, or Owlbear Rodeo? QuickSheet isn't a
> replacement — it's the always-visible scratch grid *beside* your VTT. Think DM screen,
> not battle map.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/dm-dashboard.csv
```

Edit cells live during the session. Autosaves every 5 seconds — nothing lost if you
accidentally close a window. Add to startup and your GM screen is there every boot.

## What goes on the wallpaper

The starter sheet at `examples/dm-dashboard.csv` gives you a session-ready layout:

| Column           | What it shows                                    | How                                      |
|------------------|--------------------------------------------------|------------------------------------------|
| **Initiative**   | Turn order — names sorted by roll                | [`init`](https://github.com/cemheren/quicksheet-init-ext) extension |
| **Dice**         | Quick rolls: d20, 2d6, d100, advantage           | [`roll`](https://github.com/cemheren/quicksheet-roll-ext) extension |
| **HP tracker**   | Current / max HP for each NPC or monster         | Plain cells with Σ column sum            |
| **Loot table**   | Random item from a weighted list                 | `roll: 1d12` mapped to your table        |
| **Session notes**| Quick shorthand — "Goblin fled north", timestamps| Plain cells, searchable with Ctrl+F      |

## Extensions for the table

### `roll:` — Dice roller

```
roll: 1d20+5
roll: 4d6kh3
roll: 2d10
```

Supports standard dice notation including keep-highest (`kh`), keep-lowest (`kl`),
and arbitrary modifiers. Results update on each cell refresh.

### `init:` — Initiative tracker

```
init: Gandalf 18, Goblin 12, Frodo 7, Balrog 22
```

Sorts combatants by initiative roll and displays the ordered list. Update the cell
to re-sort when new combatants join or leave.

## Recipes

### Recipe: Session-start initiative

1. Before combat, roll initiative for each NPC: `roll: 1d20+2` per creature.
2. Collect player rolls verbally.
3. Type into the `init:` cell: `init: Rogue 19, Fighter 14, Goblin1 12, Goblin2 8`.
4. The cell shows sorted turn order. Update as creatures drop.

### Recipe: Random encounter table

Set up a column where each row is an encounter. Use `roll: 1d6` in a header cell to
pick which row is "active" this turn. Pair with sticky notes in adjacent cells for
monster stats.

### Recipe: Persistent campaign tracker

Keep a second CSV (`campaign-log.csv`) with columns for Session #, Date, Location,
NPCs met, Loot gained. Switch between sheets or run two QuickSheet instances on
different monitors.

## Make it pretty

Cycle themes with **Ctrl+T** — Dracula and Synthwave fit the fantasy-noir aesthetic;
Nord for a cleaner look. The grid fades behind your VTT or PDF reader, always one
`Win+D` away.

## Why wallpaper > a second app

- **Always visible** — glance past your VTT window to see initiative order. No switching.
- **Survives reboots** — your session prep CSV loads on boot.
- **Zero internet** — works offline at the game store, the café, the cabin.
- **Editable mid-combat** — click a cell, type, done. No form dialogs.
- **Plain CSV** — back up your campaign data with git, Dropbox, anything.

## Share your setup

Post your GM-screen wallpaper on r/rpg, r/DMAcademy, r/unixporn, or r/dndnext.
If you build a custom extension for your system (Fate dice, PbtA moves, etc.),
open an issue — we'll link it in the directory.
