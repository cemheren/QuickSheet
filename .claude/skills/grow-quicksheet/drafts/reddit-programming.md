# Reddit r/programming + r/linux drafts — QuickSheet

Drafted 2026-05-14. r/programming is the largest audience but has the strictest "no self-promo" enforcement. r/linux is more forgiving of project announcements but rewards Linux-specific angles.

---

## r/programming (3.6M+ subscribers, strict rules)

r/programming is allergic to anything reading as a marketing post. Lead with the technique, not the project. Phrase as a write-up of one specific thing you learned, with the project as the artifact.

Mods usually allow:
- Technical write-ups about a specific implementation
- Posts that "teach" something
- Open-source release announcements that have meaningful technical depth

Mods usually remove:
- "I built X, please check it out"
- Pure product announcements
- Self-promotional flair

### Title (recommended)

**Embedding a TUI app as the actual desktop wallpaper on Windows (WorkerW) and Linux (X11)**

Alt titles, only if the recommended one feels too long or specific:
- `Two ways to make a terminal app render as your desktop wallpaper`
- `Reading explorer.exe's secret WorkerW, then doing the same thing in X11 ten lines later`

### Body

```
I've been wiring up a side project (QuickSheet — a TUI spreadsheet that doubles as the interactive desktop wallpaper) and wrote up the cross-platform "wallpaper" trick for anyone who has not implemented it before. The two OSes need very different code, but the underlying idea is the same: you tell the window manager that your window isn't a normal window.

**Windows (WorkerW):**

explorer.exe owns the desktop. Sending a magic SendMessageTimeout to its `Progman` window with wParam = 0x052C triggers a fork: explorer spawns a child window named `WorkerW`, behind the desktop icons but in front of the wallpaper-as-image. If you SetParent your own window into that WorkerW, your window becomes the wallpaper.

```csharp
[DllImport("user32.dll")]
static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam,
    IntPtr lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

IntPtr progman = FindWindow("Progman", null);
SendMessageTimeout(progman, 0x052C, IntPtr.Zero, IntPtr.Zero, 0, 1000, out _);
// Walk the WorkerW windows; SetParent your form into the spawned one.
```

Two edge cases you have to handle:
- explorer re-asserts z-order periodically; you need a `WndProc` hook that re-bottoms.
- Win+D ("Show Desktop") hides your window along with everything else; you detect the broadcast and re-show.

**Linux (X11):**

X11 has an explicit window type hint: `_NET_WM_WINDOW_TYPE_DESKTOP`. Set it on your top-level window via `XChangeProperty`, and the window manager treats your window like the wallpaper.

```csharp
[DllImport("libX11.so.6")]
static extern int XChangeProperty(IntPtr display, IntPtr window,
    IntPtr property, IntPtr type, int format, int mode, IntPtr data, int nelements);
```

No reparenting, no message tricks, no z-order race. Stacking WMs (Xfce, Cinnamon, GNOME-on-Xorg, KDE-on-Xorg) all honor it.

Wayland is the open problem. The X11 hint doesn't survive XWayland, and there's no equivalent freedesktop standard. `wlr-layer-shell` on wlroots-based compositors (Sway, Hyprland) would be the path forward, but I haven't shipped that yet.

Side note on why this was fun: the whole project has a hard zero-NuGet-dependencies rule, so the WorkerW + X11 + ConPTY code is all hand-written P/Invoke against system libraries. Total platform-glue code is small enough to read in one sitting.

Repo if anyone wants to see it in context: https://github.com/cemheren/QuickSheet
```

### Tips

- Lead with the technique. The repo link goes at the **bottom**.
- Code snippets in the body are mandatory for r/programming engagement.
- Mention `WorkerW` early — it's a search term that pulls in people who already know what you're talking about.
- Be ready in comments for "this trick is old / fragile / hostile" — yes, all true. Acknowledge.
- Cross-post to r/csharp 3+ days later, not same day.

### Pitfalls

- Do not include any "please star the repo" CTA. Auto-flag.
- Do not use the `Show /r/programming` flair. There isn't one; the sub is allergic to "show off" posts.
- Do not respond to top-level critique with "thanks!". Engage technically.

---

## r/linux (1M+ subscribers)

r/linux audience is less hostile to project announcements but rewards posts that are specifically about Linux behavior. Lead with the X11 path and the Wayland gap.

### Title

**A wallpaper that does something: built a TUI spreadsheet that uses `_NET_WM_WINDOW_TYPE_DESKTOP` to embed as the X11 wallpaper**

### Body

```
Side project — open source — that turns the desktop into an interactive grid: notes, runnable commands, live subprocess output, hyperlinks, sparklines, extensions installed by git URL. Runs as a normal TUI inside any terminal; runs as the wallpaper when you pass `--desktop`.

For Linux specifically, the wallpaper-embedding is a single X11 window with `_NET_WM_WINDOW_TYPE_DESKTOP` set via `XChangeProperty`. No compositor-specific quirks: it works on Xfce, Cinnamon, KDE-on-Xorg, GNOME-on-Xorg. The whole X11 glue is hand-written P/Invoke from C# — no GTK, no Avalonia, no Tk. The project has a zero-NuGet-dependencies rule, so everything's BCL + `libX11.so.6` + `libXft.so.2`.

Wayland is the open question. The X11 hint doesn't survive XWayland on any compositor I've tested. `wlr-layer-shell` (Sway, Hyprland) would be the obvious path forward — if anyone here has tried that from a non-Wayland-native language, I'd love to hear what was painful. There's an open issue for it (#3 on the repo).

Repo: https://github.com/cemheren/QuickSheet
Tour: docs/tour.md in the repo
```

### Tips

- Lead with the technique that's Linux-specific.
- Mention `XWayland` and `wlr-layer-shell` by name — these are the search terms Linux folks recognize.
- Be ready for "why .NET?" — answer: because the Windows side needs P/Invoke, .NET is the path of least resistance; the project happens to also run on Linux because conditional compilation is real.
- Be ready for "Wayland doesn't honor that hint" — yes, that's the open issue, link to #3.

### Pitfalls

- Avoid `Show HN` or "Show /r/linux" phrasing — r/linux doesn't have a "Show" custom.
- Do not double-post with r/commandline within a week.
- If your post gets ~30 upvotes in the first hour, an X11 / Wayland tribalism flame war is likely. Stay engaged on the technical questions, ignore the rest.

---

## Coordination notes

- **HN first**, then r/commandline, then r/dotnet, then r/programming, then r/linux. Stagger by 2–5 days each. Same project, different framings.
- All of them link to https://github.com/cemheren/QuickSheet. Posts that mention previous posts directly get flagged on most subs.
- After r/programming, comments will be the highest-quality technical feedback. Save the harshest critiques as issues if they expose real flaws.
