# quicksheet-init-ext

Combat initiative tracker for QuickSheet — sort combatants by initiative roll, cycle through turn order, track rounds. Built for tabletop RPG GMs who use their desktop as a behind-the-screen dashboard.

## Install

In any QuickSheet cell:

```
ext: github:cemheren/quicksheet-init-ext
```

## Usage

| Cell contents              | Result                                          |
|----------------------------|-------------------------------------------------|
| `init: add Goblin 15`     | Add "Goblin" at initiative 15, show order       |
| `init: add Ranger 18`     | Add "Ranger" at 18, re-sort and show            |
| `init: Wizard 12`         | Shorthand — same as `init: add Wizard 12`       |
| `init: show`              | Display sorted initiative order                 |
| `init: next`              | Advance to next combatant (wraps → new round)   |
| `init: prev`              | Go back one turn                                |
| `init: remove Goblin`     | Remove a combatant from the tracker             |
| `init: clear`             | Reset the entire tracker                        |
| `init: round`             | Show the current round number                   |

## How it works

- Combatants are sorted **descending** by initiative (highest goes first — standard D&D/Pathfinder convention).
- A `►` marker shows whose turn it is.
- `next` advances the marker; when it wraps past the last combatant, the round counter increments.
- Re-adding a name overwrites its previous initiative (re-roll scenario).
- State lives in-memory for the extension session. Restarting QuickSheet resets the tracker.

## Example output

```
R1: | ► 18 Ranger | 15 Goblin | 12 Wizard
```

After `init: next`:

```
R1: |  18 Ranger | ► 15 Goblin | 12 Wizard
```

## Pairs well with

- [`quicksheet-dice`](https://github.com/cemheren/quicksheet-dice) — roll initiative in one cell, then `init: add <name> <result>` in another.
- Use several `init: add ...` cells at the top of your DM screen, and one `init: show` cell that always displays the current order.

## Build

```bash
dotnet build InitExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines:

1. On startup, emit `{"type":"register","prefix":"init","name":"Initiative Tracker","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["<command>"]}`, process the command and reply with `{"type":"write","id":"...","cells":[{"r":0,"c":0,"v":"..."}]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
