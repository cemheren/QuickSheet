# Terminal Directory Submissions — QuickSheet

Drafted 2026-06-15. User submits manually via each directory's form/email.

---

## 1. TerminalTrove — https://terminaltrove.com/post/

Submit via their "Post a Tool" form at https://terminaltrove.com/post/

### Form fields

- **Tool Name:** QuickSheet
- **Website/URL:** https://github.com/cemheren/QuickSheet
- **Source:** https://github.com/cemheren/QuickSheet
- **Language:** C# (.NET 9)
- **Categories:** linux, windows, terminal, utilities, text-processing, spreadsheet, productivity, financial
- **Short description:** A terminal spreadsheet that doubles as your desktop wallpaper.
- **Full description:**

> QuickSheet is an interactive terminal spreadsheet (C# / .NET 9) that also
> embeds itself as a transparent, editable desktop wallpaper on Windows and
> Linux (X11). Cell prefixes enable runnable commands (`r: cmd`), live
> subprocess output (`i: cmd`), hyperlinks, sparklines, and inline references.
> An extension protocol (`ext: github:user/repo`) connects 50+ community
> extensions for weather, stocks, DNS, Docker status, GitHub streaks, and more.
>
> Zero NuGet dependencies — all native interop is hand-written P/Invoke.
> Persistence is plain CSV. Cross-platform: Windows (WinForms + WorkerW
> embedding) and Linux (raw X11 via libX11/libXft).

- **Install:**

```sh
# Clone and build (requires .NET 9 SDK)
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --project ExcelConsole.csproj
```

### Why it fits

TerminalTrove already lists VisiData (terminal spreadsheet). QuickSheet
differentiates with: (1) desktop-wallpaper mode as unique UX, (2) live
subprocess cells, (3) 50+ extension ecosystem, (4) zero dependencies.
The "newly added" section rotates weekly — good for a burst of visibility.

---

## 2. Console.dev — https://console.dev/tools/submit/

Submit via their web form at https://console.dev/tools/submit/

### Form fields

- **Tool Name:** QuickSheet
- **URL:** https://github.com/cemheren/QuickSheet
- **Short Description:** Terminal spreadsheet that doubles as your desktop wallpaper. Zero dependencies, CSV persistence, 50+ extensions.
- **Longer Description:**

> QuickSheet is an interactive, zero-dependency spreadsheet for the terminal
> that also runs as a transparent desktop wallpaper on Windows and Linux.
>
> Key features:
> - Cell prefixes: `r: cmd` (run command), `i: cmd` (live process output),
>   `s: 1,3,7,2` (sparklines), URLs auto-detected as clickable hyperlinks.
> - Extension protocol: `ext: github:user/repo` loads community extensions
>   (50+ available: weather, stock tickers, DNS, Docker, GitHub streaks,
>   LeetCode stats, mortgage calculators, and more).
> - Desktop mode: embeds behind icons as a live, editable wallpaper
>   (Win32 WorkerW on Windows, X11 _NET_WM_WINDOW_TYPE_DESKTOP on Linux).
> - Pure CSV persistence, autosave, Markdown/SQL export.
> - Written in C# / .NET 9 with zero NuGet packages — all platform interop
>   is hand-written P/Invoke.
>
> Think VisiData meets Rainmeter, built for developers who live in the terminal.

- **Categories/Tags:** Productivity, Terminal, Spreadsheet, Developer Tools, Desktop
- **Developer:** cemheren
- **Contact Email:** cemheren@gmail.com

### Why it fits

Console.dev features "interesting tools for developers" — weekly newsletter
with high developer audience overlap. The desktop-wallpaper angle is novel
enough to stand out. They've featured TUI tools before (lazygit, charm tools).

---

## 3. Charm CLI Newsletter mention (stretch goal)

Charm runs a newsletter + community around terminal aesthetics. Not a formal
directory, but worth an email pitch:

**To:** info@charm.sh (or their community Discord)
**Subject:** QuickSheet — terminal spreadsheet that lives on your desktop

> Hey Charm team — thought you might find QuickSheet interesting for a
> newsletter mention. It's a .NET terminal spreadsheet with a unique twist:
> it can embed itself as a transparent desktop wallpaper (X11 on Linux,
> WorkerW on Windows). Cells support live subprocess output, sparklines,
> and there's a JSON-lines extension protocol with 50+ community plugins.
>
> Zero dependencies, MIT licensed, actively developed.
> https://github.com/cemheren/QuickSheet
>
> Cheers!

---

## Submission checklist (for user)

- [ ] TerminalTrove: fill form at /post/ with above fields
- [ ] Console.dev: fill form at /tools/submit/ with above fields
- [ ] (Optional) Charm newsletter pitch email
- [ ] After listing goes live, link back from QuickSheet README ("As seen on...")
