# QuickSheet for Dungeon Masters

Your desktop wallpaper becomes the DM screen. Initiative tracker, dice roller,
NPC stats, loot tables, condition reminders — all visible behind your browser or
VTT window without alt-tabbing. No Roll20 subscription. No Foundry server. Just
a CSV on your machine.

> Already using a VTT? QuickSheet isn't a replacement — it's the GM-side scratchpad
> that sits behind every window and never needs a tab switch.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/dm-dashboard.csv
```

Edit cells live, autosaves every 5 seconds. Your initiative order, NPC hit points,
and session notes persist across reboots.

## What goes on the wallpaper

The starter sheet at `examples/dm-dashboard.csv` covers five zones:

| Zone               | What it shows                                     | How                                                              |
|--------------------|---------------------------------------------------|------------------------------------------------------------------|
| **Initiative**     | Sorted turn order with round counter              | [`init`](https://github.com/cemheren/quicksheet-init-ext)        |
| **Dice**           | Roll any expression: 1d20+5, 4d6kh3, advantage   | [`roll`](https://github.com/cemheren/quicksheet-roll-ext)        |
| **NPC Tracker**    | Name, HP, AC, status — edit mid-combat            | Plain cells                                                      |
| **Loot Table**     | Item, value, rarity — hand out on the fly         | Plain cells                                                      |
| **Session Notes**  | Campaign name, quest, location, quick references  | Plain cells                                                      |

Extensions install with a single `ext: github:cemheren/quicksheet-<name>` cell —
QuickSheet clones the repo and the prefix is live instantly.

## Initiative tracking

The [`init`](https://github.com/cemheren/quicksheet-init-ext) extension manages
combat turn order:

```
ext: github:cemheren/quicksheet-init-ext
init: add Fighter 20
init: add Ranger 18
init: add Goblin 15
init: add Wizard 12
init: show
```

Output: `R1: | ► 20 Fighter | 18 Ranger | 15 Goblin | 12 Wizard`

Use `init: next` to advance the turn marker. When it wraps, the round counter
ticks up. `init: remove Goblin` when they drop to 0 HP.

## Dice rolling

The [`roll`](https://github.com/cemheren/quicksheet-roll-ext) extension handles
any standard polyhedral expression:

```
ext: github:cemheren/quicksheet-roll-ext
roll: 1d20+5
roll: 2d6+3
roll: 4d6kh3
roll: 1d100
```

| Expression      | Meaning                            |
|-----------------|------------------------------------|
| `1d20+5`        | Roll d20 with +5 modifier          |
| `4d6kh3`        | Roll 4d6, keep highest 3 (stats)   |
| `2d8+2d6`       | Multiple dice groups, summed       |
| `1d100`         | Percentile roll for encounter table |

Each cell re-rolls on activation — click the cell to roll again.

## NPC tracker (plain cells)

No extension needed. Just type directly into the grid:

```
Name,HP,AC,Status
Goblin Chief,45,16,Alive
Shadow Drake,38,14,Bloodied
Guard Captain,52,18,Hostile
```

Edit HP in real time during combat. Ctrl+F to find an NPC by name. Autosave means
you never lose mid-session data if your laptop dies.

## Session notes and quick references

Keep a "DC reference" and "XP thresholds" block in a corner of the grid so you
never have to flip through the DMG:

```
DC Easy: 10
DC Medium: 15
DC Hard: 20
DC Very Hard: 25
DC Nearly Impossible: 30
```

Track quest info, party location, and rewards in plain cells. When the session
ends, the CSV is your session log — commit it to git for a campaign journal.

## Multi-session campaign journal

Each session's CSV is a snapshot. A workflow that works well:

```bash
cp examples/dm-dashboard.csv sessions/session-014.csv
# Edit live during session 14...
# After session, save a copy:
cp sessions/session-014.csv sessions/session-015.csv
# Start next session with a clean slate of initiative but same NPCs/notes
```

Or just keep one file and let autosave handle persistence.

## Pair with other extensions

| Extension                          | Use case                           |
|------------------------------------|------------------------------------|
| [`define`](https://github.com/cemheren/quicksheet-define-ext) | Look up a word's meaning mid-session |
| [`health`](https://github.com/cemheren/quicksheet-health-ext) | Monitor your VTT server uptime     |
| [`ping`](https://github.com/cemheren/quicksheet-ping-ext)     | Check if the Discord bot is alive  |

## Why wallpaper > app

- **Always visible** — the grid sits behind your VTT window; glance down to check initiative without alt-tab.
- **Zero cloud** — your campaign data is a CSV. No account, no subscription, no internet required during play.
- **Hackable** — write a custom extension for your homebrew system in 50 lines of any language ([extension-protocol.md](extension-protocol.md)).
- **Survives crashes** — autosaves every 5 seconds. Laptop overheats mid-boss-fight? Data is safe.

## See also

- [Full extension directory](extensions.md)
- [Keyboard shortcuts](keyboard-shortcuts.md)
- [Extension protocol spec](extension-protocol.md) — write a custom encounter-table extension in an afternoon
- [for-students.md](for-students.md) — if you're also a student running campaigns between lectures
