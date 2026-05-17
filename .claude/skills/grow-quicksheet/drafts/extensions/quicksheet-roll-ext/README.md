# quicksheet-roll-ext

A QuickSheet extension that rolls dice in a cell and (optionally) looks the result up in an encounter / loot table file. Aimed at tabletop RPG GMs running QuickSheet on their behind-the-laptop monitor while their players see only the shared map.

```
ext: github:cemheren/quicksheet-roll-ext
roll: 1d20
roll: 4d6 drop lowest
roll: 2d6+3
roll: adv 1d20
roll: dis 1d20
roll: 1d100, encounters.txt
```

## Dice expressions

| Expression                | Meaning                                            |
|---------------------------|----------------------------------------------------|
| `NdM`                     | Roll N dice with M faces, sum.                     |
| `NdM+K` / `NdM-K`         | With a +K / -K modifier.                           |
| `NdM drop lowest`         | Drop the lowest die before summing.                |
| `NdM drop highest`        | Drop the highest die before summing.               |
| `adv NdM` / `dis NdM`     | Roll twice; take the higher (advantage) or lower.  |

## Table lookup

Pass a second argument — a file path — and the rolled total is looked up in a range-prefixed table:

```
# encounters.txt
1-5     Goblins ambush from the trees
6-20    Old shrine to a forgotten god
21-40   Wagon tracks heading west
41-60   A merchant looking for an escort
61-80   Sudden thunderstorm; visibility drops
81-100  A dragon's distant roar
```

Then `roll: 1d100, encounters.txt` rolls, finds the matching row, and renders e.g. `73  →  Sudden thunderstorm; visibility drops`.

Tables use whatever path is convenient — relative to the working directory of QuickSheet, or absolute. Blank lines and lines starting with `#` are ignored.

## Why this lives on a wallpaper

- The grid is behind the laptop lid; players don't see it.
- One cell per encounter table (combat, weather, loot, NPC names, town gossip…). One click rolls.
- The table files are plain text — version-control them, share them, paste them into Obsidian.
- No Roll20, no Foundry, no account, no SaaS.

## Honesty

- **No persistent state.** Each roll is independent; there's no "history" cell. You can paste the result somewhere if you want history.
- **No Discord / VTT integration.** This is a GM-side personal screen, not a shared rolling surface.
- **No exploding dice / Fudge / dice pools.** Standard polyhedral with simple modifiers + drop-lowest/highest + adv/dis. PRs welcome for system-specific dice variants.

## Build

```bash
dotnet build RollExtension.csproj
```

.NET 9, MIT, zero NuGet dependencies.

## Protocol

Standard QuickSheet JSON-lines:

1. On startup, emit `{"type":"register","prefix":"roll","name":"Dice Roller","version":"1.0.0"}`.
2. On each `{"type":"activate","id":"...","params":["<expr>[, table.txt]"]}`, parse the expression, roll, optionally look up in the table, and reply with `{"type":"write","id":"...","cells":[{"r":0,"c":0,"v":"= 14  [4, 6, 4]"}]}`.

See [docs/extension-protocol.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extension-protocol.md).

## License

MIT.
