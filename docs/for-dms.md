# QuickSheet for tabletop GMs

If you run D&D, Pathfinder, Call of Cthulhu, or any tabletop RPG and you have a
laptop or second monitor at the table — **your wallpaper can be the DM screen.**
Initiative order, dice rolls, NPC stats, session notes, all painted *behind*
your VTT or PDF window. No browser tab, no app to alt-tab into.

> Already using D&D Beyond, Foundry VTT, or a physical DM screen? Think of this
> as the layer underneath — the at-a-glance reference sheet that's visible every
> time you peek at the desktop.

## Five-minute setup

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj -- --desktop examples/dm-screen.csv
```

Edit cells, autosaves to CSV every 5 seconds. The grid sits *behind* every
window — click through to take a note, roll dice, or check initiative, then
click back to your VTT.

## What goes on the wallpaper

Starter sheet at `examples/dm-screen.csv`. The zones:

| Zone              | What it shows                                             | How                                                                                  |
|-------------------|-----------------------------------------------------------|--------------------------------------------------------------------------------------|
| **Dice roller**   | Roll any notation — `2d6+3`, `d20`, `4d6kh3`             | [`dice`](https://github.com/Deskworks/quicksheet-dice) extension                     |
| **Initiative**    | Numbered list of combatants, editable mid-encounter       | Plain cells — type names + rolls, sort mentally or re-order rows                     |
| **NPC stats**     | AC, HP, attack bonus for the monsters in this encounter   | Plain cells — copy from your module or type freehand                                 |
| **Session notes** | Running log of what happened, decisions, NPC names        | Plain cells — the grid is your notebook                                              |
| **Quick refs**    | Condition definitions, DC table, cover rules              | Plain cells with the rules you always forget                                         |
| **Launchers**     | Open your PDF module, VTT, playlist, or random-name site  | `r: xdg-open dungeon.pdf` or `r: firefox https://donjon.bin.sh`                     |

The dice extension installs with a single cell: `ext: github:Deskworks/quicksheet-dice`.

## Dice rolling in detail

Type in any cell:

```
roll: d20
```

The extension returns the result in the rows below:

```
🎲 d20
Dice: 17
Total: 17
```

Supported notations:

| Notation    | Meaning                        |
|-------------|-------------------------------|
| `d20`       | Single die                    |
| `2d6+3`     | Multiple dice + modifier      |
| `4d6kh3`    | Roll 4d6, keep highest 3      |
| `3d6kl2`    | Roll 3d6, keep lowest 2       |
| `d%`        | Percentile (d100)             |
| `4dF`       | Fudge/FATE dice (−/○/+)      |

Critical hits (natural 20) and fumbles (natural 1) on d20 are flagged with 💥 and 💀.

## Recipe: initiative tracker

1. Dedicate a column (say column A) to initiative.
2. Row 1: header — type `Initiative`.
3. Rows 2–10: type each combatant as `14 — Goblin Archer`, `18 — Paladin`, etc.
4. Re-order rows by selecting and moving cells (Ctrl+Shift+↑/↓) as creatures act
   or get knocked out.
5. Mark dead NPCs by prefixing with `~~` or clearing the cell.

No extension needed — it's just cells. The wallpaper makes it always visible
beside your battle map.

## Recipe: encounter prep panel

Before a session, fill a few columns with the encounter's stat blocks:

```
| NPC             | AC  | HP  | Attack         | Notes             |
|-----------------|-----|-----|----------------|-------------------|
| Goblin Archer   | 15  | 12  | +4 shortbow    | hides round 1     |
| Bugbear Chief   | 16  | 45  | +5 morningstar | surprised = crit  |
| Dire Wolf       | 14  | 37  | +5 bite, prone | pack tactics      |
```

During the encounter, decrement HP in-place. The wallpaper autosaves so you
never lose mid-combat state.

## Recipe: quick-reference strip

Dedicate a row or column to the rules you always look up:

```
Conditions: Prone=-adv melee, +adv ranged>5ft | Grappled=speed 0 | Stunned=auto-fail STR/DEX saves
Cover: Half=+2AC | Three-quarters=+5AC | Full=untargetable
DCs: Easy=10 | Medium=15 | Hard=20 | Very Hard=25 | Nearly Impossible=30
```

These live on your wallpaper permanently. No flipping through the PHB mid-session.

## Launchers for session tools

Prefix cells with `r:` to make them one-click launchers:

```
r: xdg-open ~/campaigns/current/module.pdf
r: firefox https://donjon.bin.sh/5e/random/#type=encounter
r: firefox https://open.spotify.com/playlist/your-ambient-playlist
r: firefox https://www.improved-initiative.com
```

Hit Enter on the cell to fire the command. Build your whole session-start
sequence in a column and multi-select → Enter to open everything at once.

## Make it pretty

Cycle themes with **Ctrl+T**: Dark, Light, Nord, Solarized, Matrix, Dracula,
Synthwave, Gruvbox, Monokai, HotdogStand.

- **Dracula** or **Matrix** for the dark-dungeon vibe.
- **Gruvbox** for a warm parchment-ish feel.
- **Nord** if your players are in Icewind Dale.

## What this IS NOT

- **Not a VTT.** No maps, no tokens, no fog of war. Use Foundry, Roll20, or
  Owlbear Rodeo for that. QuickSheet sits beside or behind your VTT.
- **Not a character sheet.** Players manage their own sheets. The DM screen is
  for *your* reference — NPC stats, initiative, notes.
- **Not a campaign manager.** No relational database, no session trees. It's a
  CSV grid. For deep campaign management use Notion, Obsidian, or Kanka.

It's the *glanceable layer*: the stuff you need to see without switching
windows. Roll dice, check AC, note what happened, move on.

## More extensions to mix in

- [`worldtm`](https://github.com/Deskworks/quicksheet-worldtm) — world clocks,
  handy for remote sessions across time zones.
- [`pomo`](https://github.com/Deskworks/quicksheet-pomodoro) — session timer
  for structured play blocks.
- [`words`](https://github.com/cemheren/quicksheet-words) — word count for
  session writeups.
- [Full extension directory](extensions.md).

## Share your setup

Post your DM-screen wallpaper on r/DnD, r/DMAcademy, r/rpg, or r/unixporn
(the rice angle is real — a tabletop DM screen as a desktop rice is novel).
Open an issue if you want a new extension officially listed.
