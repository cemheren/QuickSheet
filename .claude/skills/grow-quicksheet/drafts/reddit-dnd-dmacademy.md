# r/DnD + r/DMAcademy submission draft — QuickSheet as a DM screen

Drafted 2026-06-14. Pairs with `docs/for-dms.md` (PR #359) and the TTRPG
extension trio: `quicksheet-dice`, `quicksheet-init-ext`, `quicksheet-cal-ext`.

> **Target subs:** r/DnD (~3.5M members), r/DMAcademy (~700k). Same core text,
> tone shifts noted below. r/DMAcademy rewards practical "here's how" content;
> r/DnD is more show-and-tell. Post in both but space them 2–3 days apart.

---

## Title options

### r/DMAcademy (practical angle)
1. `I replaced my physical DM screen with a live spreadsheet on my desktop wallpaper`
2. `My prep-to-table workflow: a wallpaper-embedded spreadsheet with dice, initiative, and random tables`
3. `Free tool that turns your desktop into a DM screen — dice, initiative tracking, random encounter tables`

### r/DnD (show-and-tell angle)
1. `My desktop wallpaper IS my DM screen now — dice roller, initiative tracker, random tables, all in one grid`
2. `Built an interactive DM screen that lives on my desktop wallpaper`
3. `[OC] My wallpaper is a live spreadsheet — I use it as a digital DM screen`

**Recommended:** r/DMAcademy #1, r/DnD #1.

---

## Post body (r/DMAcademy version — practical, tutorial-style)

**Flair:** `Resources`

---

Hey DMs,

I run games on a laptop and got tired of tabbing between initiative trackers,
random tables, and dice rollers. I wanted everything visible at a glance without
covering my VTT or notes.

I found [QuickSheet](https://github.com/cemheren/QuickSheet) — a free,
open-source spreadsheet that embeds directly into your desktop wallpaper. It's
always visible behind your windows, takes zero screen real estate, and you can
click into it anytime.

Here's what my DM screen looks like during a session:

**[SCREENSHOT PLACEHOLDER — wallpaper showing grid with dice results, initiative
order, random encounter table, session notes]**

### What makes it useful for DMing

**Dice rolling in cells:**
Type `roll: 2d6+3` in any cell and it rolls instantly. Supports all the common
patterns — `d20`, `4d6kh3` (keep highest 3), `2d8+1d6+5`, percentile. I keep a
column of pre-labeled rolls: attack, damage, wild magic, random encounter.

**Initiative tracking:**
The initiative extension (`init:`) lets you sort combatants, cycle turns, and
track round count. Type names and rolls into cells, and it handles ordering.

**Random tables:**
This is the killer feature for me. I put my random encounter tables, NPC name
lists, treasure tables, etc., directly in the grid. Combined with `roll:` I can
reference a cell range and get a random pick.

**Session notes that persist:**
Everything autosaves to a CSV file. I keep a "Session Log" column where I jot
notes during play. After the session I export to markdown for my campaign wiki.

### Setup (5 minutes)

1. Install: `dotnet run --project ExcelConsole.csproj -- --desktop`
   (needs .NET 9 runtime — free from Microsoft)
2. Install dice: type `ext: github:Deskworks/quicksheet-dice` in any cell, Enter
3. Install initiative: type `ext: github:cemheren/quicksheet-init-ext` in any cell, Enter
4. Start filling in your random tables, NPC names, encounter lists

It's completely free, no account needed, no internet required after setup. Works
on Linux and Windows. All data stays local in a plain CSV file.

### What it's NOT

- Not a VTT replacement (no maps, no tokens)
- Not a character sheet manager
- Not cloud-synced (local CSV only, which is a feature for privacy)

It's purely a DM-side tool for the stuff you'd put on a physical screen: quick
references, dice, initiative, notes.

### Links

- GitHub: https://github.com/cemheren/QuickSheet
- DM guide: https://github.com/cemheren/QuickSheet/blob/main/docs/for-dms.md
- Dice extension: https://github.com/Deskworks/quicksheet-dice
- Initiative tracker: https://github.com/cemheren/quicksheet-init-ext

Happy to answer questions about the workflow. I've been using this for ~3 months
of weekly sessions and it's genuinely sped up my combat rounds.

---

## Post body (r/DnD version — show-and-tell, shorter)

**Flair:** `OC` or `Resources`

---

My desktop wallpaper IS my DM screen now.

**[SCREENSHOT PLACEHOLDER]**

I use [QuickSheet](https://github.com/cemheren/QuickSheet) — a free open-source
spreadsheet that embeds into your wallpaper. It's always visible behind windows,
takes no screen space, and everything autosaves.

What's in my grid:
- **Dice:** `roll: 2d6+3`, `roll: d20`, `roll: 4d6kh3` — results update in-cell
- **Initiative:** sorted combatant list with turn cycling
- **Random tables:** encounter tables, NPC names, loot — one click to roll on them
- **Session notes:** jot during play, export to markdown after

Setup is free, no internet needed after install, no account. Linux + Windows.
All data stays in a local CSV.

Links:
- [QuickSheet](https://github.com/cemheren/QuickSheet)
- [DM usage guide](https://github.com/cemheren/QuickSheet/blob/main/docs/for-dms.md)
- [Dice extension](https://github.com/Deskworks/quicksheet-dice)

---

## Posting notes

1. **Screenshot is mandatory.** Without a compelling image showing the actual
   DM-screen layout, the post will die. Capture with a real game grid — random
   tables populated, dice results visible, initiative list mid-combat.
2. **Timing:** r/DMAcademy peaks weekday evenings US time (6–9pm ET). r/DnD is
   more forgiving but weekday evening still best.
3. **Engagement:** Answer every comment within 2 hours. Common questions will be:
   - "Does it work on Mac?" → Not yet, Linux + Windows only.
   - "Can players see it?" → No, it's DM-side only (wallpaper).
   - "Why not just use Google Sheets?" → Always visible, no browser tab, works
     offline, dice/initiative built in, embeds in wallpaper.
4. **Don't over-shill.** One mention of the GitHub link is enough. Let the
   screenshot and workflow description do the selling.
5. **Cross-post spacing:** Post r/DMAcademy first (higher signal audience, more
   likely to generate quality discussion). Wait 3 days, then r/DnD.
6. **r/DMAcademy rules:** no self-promotion flair — use `Resources`. The post
   must be genuinely useful, not just "look at my project." The tutorial framing
   achieves this.
