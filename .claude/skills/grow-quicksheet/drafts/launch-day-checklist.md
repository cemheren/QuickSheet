# Launch Day Checklist

**When:** A Tuesday, 6:30–9:30am EST. Pick a week when you have 2 hours free
to monitor HN comments. Ideally after merging the release workflow PR (#308)
and tagging `v1.0.0`.

**Prerequisites (must have BEFORE launch day):**
- [ ] Demo GIF captured (3–5 seconds showing: open desktop → type in cell →
      run command → see output). Place at top of README.
- [ ] `v1.0.0` release tagged with binaries (merge PR #308 first).
- [ ] 3–5 issues labeled "good first issue" (already on the repo).
- [ ] README has subtle star CTA near Quick Start.

---

## Launch sequence (execute in order, within ~90 minutes)

### 1. Hacker News — Show HN (T+0 min)

**Post title:**
```
Show HN: QuickSheet – Turn your desktop wallpaper into an interactive spreadsheet
```

**First comment** (post immediately after submitting — this is your pitch):
```
Hey HN! I built QuickSheet because I had a second monitor showing a static
wallpaper doing nothing. Now it's a transparent grid where I pin notes, launch
apps, track numbers, and run shell commands — all without opening a window.

Key decisions:
- Zero NuGet dependencies. All native interop (X11, WinForms, ConPTY) is
  hand-written P/Invoke. The entire supply chain is the .NET 9 SDK.
- Data is just a CSV file. Same file works on Windows and Linux.
- 40+ extensions via `ext: github:user/repo` — weather, stocks, system
  monitoring, AI prompts — each installs in one cell.

Works on Linux (X11, _NET_WM_WINDOW_TYPE_DESKTOP) and Windows (WinForms
WorkerW embedding). Autosaves every 5 seconds.

Would love feedback on the extension protocol or use cases I haven't thought of.
```

**Then:** Stay online for 60–90 minutes. Reply to EVERY comment quickly.
Engagement speed directly affects HN ranking algorithm.

### 2. Reddit — r/commandline (T+15 min)

Use draft from `drafts/reddit-commandline.md`. Crosspost timing: wait 15 min
after HN to avoid looking like spam.

### 3. Reddit — r/linux (T+20 min)

Use draft from `drafts/reddit-linux.md`. Emphasize the X11 wallpaper angle.

### 4. Reddit — r/unixporn (T+30 min)

**Only if you have a polished screenshot/rice.** r/unixporn requires visual
proof. Use the desktop screenshot with extensions running. Tag with your WM.

### 5. Twitter/X thread (T+45 min)

Use draft from `drafts/twitter-thread.md`. Pin the thread. Include the GIF.
Tag @climagic, @nixcraft if the tone fits.

### 6. Awesome-list PRs (T+60 min)

Submit the drafts that are ready:
- `drafts/awesome-linux-software.md` → PR to awesome-linux-software
- `drafts/awesome-windows.md` → PR to Awesome-Windows
- `drafts/awesome-csv.md` → PR to awesomeCSV

### 7. Newsletter/directory submissions (T+90 min)

- Console.dev: use `drafts/console-dev-submission.md`
- TerminalTrove: submit via their form

---

## Post-launch (days 2–7)

- Monitor HN for follow-up questions; reply daily.
- If front-paged: update README with "As seen on Hacker News" or star-history
  graph showing the spike.
- Merge the best open PRs (pick 3–5 from the 14 pending) to show active
  development.
- Thank first stargazers in a Discussion post.
- Post a "100 stars" milestone if reached — this creates secondary sharing.

---

## Failure mode

If HN doesn't front-page (likely <5 points after 1 hour):
- Don't delete. Let it sit.
- Wait 2+ weeks, then try again with a different title angle.
- Alternative title: `Show HN: I replaced my wallpaper with a CSV that runs
  shell commands`
- Consider reposting on a Wednesday if Tuesday was too competitive that week.
