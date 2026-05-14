# dev.to / Medium long-form draft — QuickSheet

Drafted 2026-05-14. ~1500 words, code-walkthrough oriented. Publishable on dev.to, Medium, or as a GitHub Pages post. Adapt the title and intro per platform.

---

## Title

**I turned my desktop wallpaper into a spreadsheet (and you can too)**

Subtitle alternatives:
- *How I built QuickSheet — a zero-NuGet .NET 9 TUI that doubles as a Windows + Linux desktop wallpaper*
- *Embedding a terminal app as the desktop wallpaper, on two operating systems, without taking a single dependency*

## Cover image

The launcher screenshot (`image-4.png` in the repo). Caption: "Cells holding runnable commands, hyperlinks, and live subprocess output. Running as the desktop wallpaper behind every window."

## Tags (dev.to)

`dotnet`, `csharp`, `terminal`, `linux`, `windows`, `showdev`

---

## Body

### The itch

I never interact with my wallpaper. It's a screensaver with delusions of grandeur — a static image I'll see between window switches, but never click. Meanwhile I always have a few things I want at hand: a couple of notes I'm jotting, the URLs I open every morning, app launchers I haven't put in a real launcher, and small running totals (weights, miles, dollars) that don't deserve a database.

I wanted those things on the wallpaper itself.

**QuickSheet** is what I ended up building. It's an interactive spreadsheet that runs as a normal TUI in any terminal — or, with `--desktop`, embeds itself as the wallpaper. Same C# code, same CSV file. Cells hold text by default, but a prefix flips them into something more interesting:

```
r: code .              ← runs the command on Enter
i: ping example.com    ← live subprocess output streams into the cell
s: 4,7,9,3,8,12       ← renders as ▂▃▆▁▅█ sparkline
L: A10, 5m            ← loops the target cell every 5 minutes
ext: github:user/repo  ← installs an extension, registering a new prefix
https://news.ycombinator.com  ← auto-detected hyperlink, opens on Enter
```

If you've ever wired together `r:` + `i:` + `L:` and watched a row of cells become a tiny live dashboard, you'll get why this is fun. If you haven't, here's a 60-second tour: [docs/tour.md](https://github.com/cemheren/QuickSheet/blob/main/docs/tour.md).

This post is about how QuickSheet works — three things I think are worth writing about:

1. The "wallpaper trick" on Windows and Linux.
2. Why I refused to take a single NuGet dependency.
3. Extensions that are just git URLs.

### 1. Making a TUI the actual desktop wallpaper

On Windows, the technique is twenty years old but still works. Explorer.exe spawns a hidden window called `WorkerW` behind the desktop icons; you can `SetParent` your own window into it, and that puts it behind everything but in front of the wallpaper-as-image. The magic incantation is a `SendMessageTimeout` to `Progman` with `0x052C` as `wParam` — that's the message that tells Explorer to spawn the `WorkerW`.

Hand-written P/Invoke, no library:

```csharp
[DllImport("user32.dll")]
private static extern IntPtr FindWindow(string lpClassName, string? lpWindowName);

[DllImport("user32.dll")]
private static extern IntPtr SendMessageTimeout(
    IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam,
    uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

IntPtr progman = FindWindow("Progman", null);
SendMessageTimeout(progman, 0x052C, IntPtr.Zero, IntPtr.Zero,
                   0, 1000, out _);
// Find the spawned WorkerW, SetParent your form into it.
```

There are two edge cases you have to handle:

- **Z-order**: explorer.exe periodically re-asserts itself. You need a `WndProc` hook that re-bottoms your window when this happens.
- **Win+D ("Show Desktop")**: it hides every window — including yours. You either detect the broadcast and re-show, or you accept that Show Desktop kills the wallpaper grid for that session. QuickSheet detects and recovers.

On Linux it's much simpler. X11 has an explicit window hint:

```
_NET_WM_WINDOW_TYPE_DESKTOP
```

Set it on your top-level window and the window manager treats it like the wallpaper. No reparenting, no message tricks. Stacking WMs (Xfce, Cinnamon, GNOME-on-Xorg, KDE-on-Xorg) all honor it. Wayland is the open problem — most compositors don't honor the X11 hint via XWayland, and there's no equivalent freedesktop standard. (Layer-shell on wlroots-based compositors would be the path forward — see issue #3.)

Again, raw P/Invoke:

```csharp
[DllImport("libX11.so.6")]
private static extern int XChangeProperty(
    IntPtr display, IntPtr window, IntPtr property, IntPtr type,
    int format, int mode, IntPtr data, int nelements);
```

Total Linux interop code: a couple hundred lines of `[DllImport]` against `libX11.so.6` and `libXft.so.2`. No GTK, no Avalonia, no Tk.

### 2. Why no NuGet dependencies

This is the constraint I'm most happy about, and the one that keeps surprising me.

The `.csproj` has zero `<PackageReference>` entries. Everything QuickSheet does — Win32 window embedding, ConPTY for live subprocess cells on Windows, X11 + Xft rendering on Linux, CSV parsing, the extension protocol — is built on what ships with .NET 9.

Three concrete reasons:

1. **Supply chain.** A clone-and-run project should not require a stranger's NuGet package to start. The fewer hands a binary has passed through, the easier I can trust it on my own desktop.
2. **The constraint forces honesty.** When you can't `dotnet add package` your way out of a problem, you have to read a docs page (or a header file) and write five lines of P/Invoke. Most of the time those five lines are the thing you actually wanted — the rest of the package was scaffolding to make it ergonomic.
3. **Cross-platform via conditional compilation, not abstraction layers.** Each OS gets its own folder under `Platform/`, the wrong one is `<Compile Remove>`d in the csproj, and shared files use `#if PLATFORM_WINDOWS / #elif PLATFORM_LINUX`. No DI container, no plugin loader, no "platform service" interface. It's startlingly readable.

I want to be honest about what this costs. Some things take longer to build this way. The ConPTY plumbing on Windows took an afternoon — `Spectre.Console` would have given me a child-process surface for free. The X11 font rendering would have been a one-line `Avalonia.Controls.TextBlock`. But once those are written, they're tiny, transparent, and yours.

### 3. Extensions are git URLs

QuickSheet has an extension system, but it doesn't have a registry, a package manager, a marketplace, a manifest format, or a sandbox. An extension is a git repo that ships:

- a `quicksheet-extension.json` manifest with a `prefix` and an `entry` command,
- a subprocess that talks JSON-lines on stdin/stdout.

When you type `ext: github:cemheren/quicksheet-weather` into a cell, QuickSheet clones the repo, reads the manifest, runs the `entry` command, sends `{"type":"init"}`, gets back `{"type":"register","prefix":"wthr",...}`, and now `wthr:` is a live cell prefix. That's it.

The protocol is two message types — `init` and `activate`. Extensions can be any language: the reference set is .NET 9 (because that's what the host is), but Python stdlib or Go would work equally well. I've shipped a handful so far:

- AI in a cell (Copilot)
- 7-day weather forecast
- TLS certificate checker (expiry, issuer)
- Crypto price quotes (CoinGecko)
- Inline dictionary
- Mortgage calculator
- Pomodoro timer with progress bar
- MX record lookup (DNS-over-HTTPS)
- HTTP probe (status + latency)

Each is a separate repo, MIT, zero NuGet, ~150 lines. The reason this works is that the *protocol* is small enough that an extension *is* its implementation — there's no framework to learn. If you've used cell prefixes for an afternoon, you understand the entire model. The full directory is at [docs/extensions.md](https://github.com/cemheren/QuickSheet/blob/main/docs/extensions.md).

### What this is good for

Not Excel. Not a database. Not a productivity app for non-technical people.

QuickSheet is for the developer who already has a tmux + bash + Raycast + Rainmeter + sticky-notes setup and would prefer one CSV-backed surface instead. It's for the SRE who wants a passive status display on the wallpaper. It's for me. Maybe it's for you.

If it sounds interesting, the repo has a 60-second tour, six concrete dashboard recipes, and screenshots:

➡ [github.com/cemheren/QuickSheet](https://github.com/cemheren/QuickSheet)

If it doesn't sound interesting — also fine. Side projects are allowed to have a small audience.

### What I'd love feedback on

- Wayland support. The XWayland path is broken in a fundamental way and I haven't solved it. If you've worked with `wlr-layer-shell` from a non-Wayland-native language, I want to hear what was painful.
- The cell-prefix concept. Is this a sane way to extend a TUI? Or am I reinventing Emacs poorly?
- The zero-NuGet rule. When does it actually hurt vs. when does it just feel virtuous?

Hit me up in the issues or wherever you find this post.

---

## Publishing notes

- **dev.to**: prepend `cover_image:` frontmatter with a hosted screenshot URL. Tags inline.
- **Medium**: paste the body, add 2–3 highlighted callouts. Medium's algorithm rewards 7+ minute read estimates — this is ~6 minutes; consider adding one more code block to push it over.
- **Personal blog / GitHub Pages**: ideal home; no platform downsides.
- **Cross-post timing**: dev.to first, then Medium 5–7 days later with a "originally posted at..." note. Don't double-post within 24h — both platforms downrank.
