# r/freelance + r/consulting post — QuickSheet freelancer dashboard

Drafted 2026-06-15. Target subs: r/freelance (~600k), r/consulting (~200k), r/freelanceWriters (~150k).

> **Position in launch sequence:** After the for-freelancers landing page (PR #364) merges and the user has a live screenshot. This is a "look what I built for my workflow" post, not a project announcement.

## Framing rule

Both subs reward personal-workflow stories ("here's how I solved X") over product announcements. Lead with the *problem*, show the *result* in a screenshot, mention the tool casually. Never say "I built this app" in the title — say what the app *does for you*.

## r/freelance version

### Title

```
I replaced my desktop wallpaper with an always-visible invoice tracker + hours log
```

### Body

```
Solo dev here. I was sick of tabbing between Notion/Sheets/email to check
which invoices are overdue and how many hours I'd logged this week. I wanted
something that's *always there* without opening a window.

I turned my desktop wallpaper into a transparent grid that shows:

- **Invoice tracker** — client, amount, days outstanding, red/green status
- **Weekly hours** — Mon-Sat with a running column sum
- **Quarterly tax countdown** — days until the next estimated payment,
  calculated from a 1099 extension
- **Expense sparklines** — YTD spend by category with mini trend lines
- **Mileage log** — IRS standard rate × miles driven, auto-calculated

It autosaves every 5 seconds to a CSV. I click any cell on the desktop to
edit — no app to open, no window to find. The data file is just a CSV so
I can `cat` it in scripts, pipe it, or open it in Excel if I need to.

[screenshot placeholder]

The tool is QuickSheet (github.com/cemheren/QuickSheet) — .NET, zero
dependencies, runs on Linux and Windows. There's a `docs/for-freelancers.md`
with a ready-to-use CSV template if anyone wants to try the same layout.

Happy to answer questions about the setup. The killer feature for me is
honestly just the autosave + always-visible part — I stopped forgetting
to follow up on 30-day invoices.
```

### Flair

r/freelance: "Tips & Advice" or "Discussion" (varies by sub rules)

---

## r/consulting version

### Title

```
Built myself an always-on project dashboard that lives in the desktop wallpaper
```

### Body

```
Independent consultant. I juggle 3-4 clients at once and kept losing track
of billable hours, outstanding invoices, and quarterly estimated tax
deadlines in the noise of tabs and apps.

My solution: I replaced my wallpaper with a transparent spreadsheet grid
that's always visible behind my windows. One glance shows me:

- Client invoices and aging (red if overdue)
- This week's billable hours by day with auto-sum
- Quarterly estimated payment countdown (Sep 15 = 92 days)
- YTD expenses with sparkline trends

Click a cell to edit. Autosaves every 5s to a plain CSV file. I start my
day by glancing at the desktop instead of opening three different tools.

It's an open-source tool called QuickSheet — zero dependencies, just
clone and build. Works on Windows and Linux. The CSV is portable so I
can version-control it or pipe it into reports.

[screenshot placeholder]

Anyone else using unconventional tools for client/project tracking? Curious
what other solos have cobbled together.
```

### Flair

r/consulting: "Discussion" or "Tools"

---

## r/freelanceWriters version (optional, lighter touch)

### Title

```
Using my desktop wallpaper as a pitch tracker / invoice board (no apps to open)
```

### Body

```
I write for 4-5 clients. I needed something *always visible* — not another
tab to forget about. I set up my desktop wallpaper as a transparent grid
where each row is a pitch or invoice:

- Pitch status: sent / accepted / published / paid
- Invoice aging: green if <30 days, red if overdue
- Weekly word count with a running total

It's just a CSV file that autosaves. Click any cell to edit. The tool is
free and open-source (QuickSheet on GitHub). Not affiliated, just a user
who found it fits the "freelancer needs a glanceable dashboard" niche
perfectly.

[screenshot]

What do you all use for tracking pitches and payments at a glance?
```

---

## Posting notes

- **Do NOT post until** the for-freelancers landing page is merged and a real
  screenshot exists. The screenshot IS the post.
- **Best times:** r/freelance peaks weekday mornings US Eastern (9-11am).
  r/consulting is similar.
- **Engagement:** Answer every comment in the first 2 hours. Be helpful about
  the workflow, not promotional about the tool.
- **Cross-post spacing:** Don't post to all three on the same day. Space 2-3
  days apart. Start with r/freelance (largest), then r/consulting, then
  r/freelanceWriters.
- **Self-promotion rules:** r/freelance allows tool mentions in context of
  workflow posts. r/consulting is stricter — frame as discussion, not showcase.
  r/freelanceWriters requires 10:1 participation ratio before self-promo.
- **Don't link directly to repo in title or first line.** Bury it naturally
  in the body after showing the value.
