# Console.dev submission — draft (2026-06-05)

> User submits manually at https://console.dev/submit. Do not post anywhere.

## Submission fields

**Tool name:** QuickSheet

**URL:** https://github.com/cemheren/QuickSheet

**Short description (1–2 sentences):**

QuickSheet replaces your desktop wallpaper with a transparent, interactive spreadsheet grid. Pin notes, launch apps, run shell commands, stream live subprocess output, and query 69+ extensions — all without opening a window.

**Longer description (2–4 sentences):**

QuickSheet is a .NET 9 desktop app that embeds a CSV-backed grid behind your windows using platform-native tricks (WorkerW on Windows, X11 `_NET_WM_WINDOW_TYPE_DESKTOP` on Linux). Cells support shell commands (`r: code .`), inline live output (`i: htop`), sparklines, hyperlinks, and a JSON-lines extension protocol with 69+ community extensions covering weather, stocks, RSS, system monitoring, devops health checks, and more. Zero NuGet dependencies — the entire supply chain is the .NET SDK. Autosaves every 5 seconds to a plain CSV file that works on both platforms.

**Category/Tags:** Desktop, CLI, Productivity, Terminal, Spreadsheet, CSV, .NET

**Key features:**

- Transparent grid lives behind all windows — click desktop to edit
- Cells run shell commands, stream subprocess output, render sparklines
- 69+ JSON-lines extensions (weather, stocks, RSS, Docker, TLS, GitHub Actions…)
- Zero dependencies — clone, `dotnet build`, run
- Cross-platform: Windows (WorkerW) + Linux (X11)
- Plain CSV persistence, autosaves every 5s
- MIT licensed, .NET 9

**Contact:** cemheren (GitHub)

## Notes for the submitter

- Console.dev's form at https://console.dev/submit is straightforward — paste the short description and URL. They may ask follow-up questions.
- Console.dev covers broader dev tools, not just terminal-only. Position as a "desktop productivity" tool for developers who live in terminals/IDEs.
- Their newsletter reaches 30k+ subscribers — heavily dev-skewed audience. A feature here could drive 20–50 stars from the long tail.
- No screenshot/image strictly required for submission (unlike Terminal Trove), but including one would help. If available, use the existing `desktop-wallpaper-commands.png` from the repo.
- If they ask what makes it different from Rainmeter/Conky: those render pre-built widgets; QuickSheet renders editable cells you type into, backed by a CSV you can version-control.
- Console.dev publishes weekly — expect 1–4 week lead time between submission and potential feature.

## After submission

- If featured: passive discovery from 30k newsletter subscribers + Console.dev archive page (indexed by search engines).
- Cross-reinforces other submissions (awesome-lists, Terminal Trove) — multiple touchpoints compound.
- Realistic outcome: 10–30 stars over several weeks if featured.
