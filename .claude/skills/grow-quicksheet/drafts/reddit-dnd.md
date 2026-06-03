# Reddit r/DnD (or r/dndnext / r/rpg) — QuickSheet as a DM Screen

Drafted 2026-06-03. Bucket C — content draft. User posts manually.

## Subreddit options (pick one)

| Sub | Fit | Notes |
|-----|-----|-------|
| r/DnD | 4.1M members | General D&D, allows tool posts on weekdays |
| r/dndnext | 1M members | 5e-focused, more mechanical audience |
| r/rpg | 1.1M members | System-agnostic, broader TTRPG audience |
| r/DMAcademy | 800K members | DM tools explicitly welcome |

**Best fit:** r/DMAcademy — they actively seek tools that solve session-prep and at-the-table friction. Less noise than r/DnD.

## Title

```
I built a desktop-wallpaper spreadsheet that doubles as a DM screen — dice rolling, initiative tracking, encounter tables, all in one grid
```

Alternative (shorter):
```
Free desktop-wallpaper app that works as a DM screen — dice, initiative, encounter tables in a live spreadsheet grid
```

## Post body (text post, not link post)

---

**The pitch:** I wanted a DM screen that lived *on my desktop* — always visible behind my windows, no alt-tabbing to a separate app. QuickSheet is a spreadsheet that renders directly on your wallpaper (Windows or Linux/X11). It has extensions for dice and initiative tracking that make it work as a lightweight DM screen.

**What it does at the table:**

- **Dice rolling** — type `roll: 2d6+3` in any cell and it evaluates. Supports `4d6kh3` (keep highest 3), advantage/disadvantage (`adv d20`), and crit detection. Results update live.
- **Initiative tracker** — `init:` extension sorts combatants by roll, tracks turn order, handles ties. You see the full combat order in a column that auto-sorts.
- **Encounter tables** — pair a dice roll with a lookup table file. `roll: d100, encounters.csv` gives you a random encounter from your prep file.
- **NPC/Monster stat blocks** — just fill cells with AC, HP, attacks. It's a spreadsheet — you already know how to lay this out.
- **Session notes** — CSV persistence means your prep file is always there next session. Autosaves every 5 seconds.

**What it looks like:** The grid sits behind your desktop icons (like Rainmeter widgets but for tabletop data). You can have your VTT/browser on top and glance at initiative order underneath.

**Extensions used:**
- `ext: github:cemheren/quicksheet-dice` — dice roller
- `ext: github:cemheren/quicksheet-init-ext` — initiative tracker

**Install (one command):**
```
# Windows (scoop) or download from GitHub Releases
# Linux
dotnet tool install -g quicksheet  # if published
# Or grab the binary from https://github.com/cemheren/QuickSheet/releases
```

**Not a VTT replacement.** This doesn't do maps, tokens, or fog of war. It's for the behind-the-screen data — initiative, HP tracking, quick rolls, session notes — that you'd normally scrawl on paper or juggle between 3 browser tabs.

**Free, open-source, zero dependencies.** MIT license. https://github.com/cemheren/QuickSheet

Happy to answer questions about setup or how to lay out a DM screen template with it.

---

## Example CSV template to include (as a follow-up comment or linked gist)

```csv
=== SESSION 12: The Sunken Vault ===,,,,
,,,,
--- INITIATIVE ---,,,,
init: Kira (Rogue),init: Thane (Paladin),init: Goblin x3,init: Wyvern,
,,,,
--- QUICK ROLLS ---,,,,
roll: d20,roll: 2d6+3,roll: d100,roll: 4d6kh3,
,,,,
--- ENCOUNTER TABLE ---,,,,
roll: d12 (forest encounters),,,
1-2: Wolves (CR 1/4 x4),,,
3-4: Bandit patrol (CR 1/8 x6),,,
5-6: Lost merchant (NPC),,,
7-8: Ruined shrine (puzzle),,,
9-10: Owlbear (CR 3),,,
11-12: Green hag in disguise (CR 3),,,
,,,,
--- MONSTER STATS ---,,,,
Name,AC,HP,Attack,Damage
Goblin,15,7,+4,1d6+2
Wyvern,13,110,+7 (bite),2d6+4
```

## Posting strategy

- **Flair:** "Resources" or "OC" depending on sub rules.
- **Time:** Tuesday–Thursday, 10am–1pm EST (peak engagement for text posts on DM subs).
- **Follow-up comment:** Post the CSV template as a code block in a reply to your own post — gives readers something to immediately try.
- **Screenshot:** If possible, include a screenshot of the grid on a dark desktop with the DM screen layout visible. This is the hook — text posts with images get 2-3x engagement on these subs.
- **Cross-post:** After 24h on r/DMAcademy, cross-post to r/DnD if it gets traction.

## Why this should work

1. **DMs are perpetually tool-hungry** — any post offering a free tool for session management gets attention.
2. **"Desktop wallpaper" angle is novel** — nobody else is doing this for TTRPG. It's visually distinctive.
3. **Dice + initiative in a spreadsheet** = immediately understandable value prop for anyone who's used a Google Sheet as a DM screen.
4. **Low bar to try** — download binary, point at a CSV. No account, no SaaS, no subscription.
