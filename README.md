# QuickSheet

**Your desktop is a spreadsheet.**

QuickSheet replaces your wallpaper with a transparent, interactive grid. Pin notes, launch apps, track tasks — all without opening a window. I wanted to give a always present feeling, there is always something lightweight in the background to achieve things. I normally find the default desktop experience pretty useless. 

![.NET 9](https://img.shields.io/badge/.NET-9.0-purple)
![License](https://img.shields.io/badge/license-MIT-blue)
![Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows)
![Linux](https://img.shields.io/badge/platform-Linux-FCC624?logo=linux&logoColor=black)

## Disclaimers

This repo contains lots of AI coding, so I can't attest for the code quality. It's meant to be a side project. I'm however taking the time to ensure it has no downstream dependencies, I wanted to be cognizant of supply chain attacks. So download locally and build locally (shouldn't need any packages for the moment.) 

I also highly recommend setting it as a startup application, this way every time there is a restart you have your notes etc. 

## Quick Start

I recommend running in release mode so things feel snappier. 

```bash
git clone https://github.com/cemheren/QuickSheet.git
cd QuickSheet
dotnet run -c Release --desktop
```

Requires [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0).

## Desktop environment for developers 
### Type directly into cells 
For quick notes, todos. Everything autosaves every 5 seconds. You can point and click to any part of your desktop to take some quick notes, use it as a buffer etc without losing focus of the application you are running. 
  
### Launcher
<!-- ![Launcher example](docs/screenshots/use-case-launcher.png) -->
Prefix any cell with `r: ` to make it a runnable command — e.g. `r: code .` or `r: firefox`.
You can open or launch multiple repos with a single operation. Multi select cells, and hit enter to run. 

I've used it to start the repos I want to work on for the day, and launch copilot with some saved prompts like summarize emails. Not sure how others solve this problem, but this to me is simpler than running startup scripts. 

![alt text](image-4.png)

### Hyper-Link Dashboard
<!-- ![Links example](docs/screenshots/use-case-links.png) -->
Paste URLs into cells. They're highlighted and open in your browser on Enter/double-click.
Similar to the launcher funcitonality you can open and run multiple by selecting multiple cells. I was going for a emacs buffer type of feel to save and run multiple cells. 

![alt text](image-1.png)

### Lightweight Data Tracking
<!-- ![Data example](docs/screenshots/use-case-data.png) -->
Auto-sum (Σ) per column and auto-product (Π) per row in the status bar. Import/export CSV. I hate opening the calculator for simple operations. This helps with that. 

![alt text](image-2.png)

<!-- 
## Add your own sections here!
Some ideas:
- Daily standups / sprint tracking
- Monitoring dashboard with r: commands
- Project-specific bookmarks
-->

### Desktop files
Desktop files are added to cells (padded to right), which can be used in a multi-select way. Helpful for finding/launching multiple files. Not sure about usability of these yet, likely I'm going to tweak this. 

![alt text](image.png)

## Extensions — Make Your Desktop Do More

QuickSheet has a lightweight extension system that lets you bring new capabilities right into your grid. Extensions are standalone programs that communicate with QuickSheet over a simple JSON-lines protocol — install one in seconds and it just works.

### Install an extension in one cell

Type `ext: github:user/repo` into any cell and press Enter. QuickSheet clones the repo, reads its manifest, and starts the extension automatically. That's it — no package managers, no config files.

### Example: GitHub Copilot on your desktop

The [quicksheet-copilot-ext](https://github.com/cemheren/quicksheet-copilot-ext) extension brings AI directly into your spreadsheet grid. Ask questions, generate structured data, or summarize cell ranges — all without leaving your desktop.

```
ext: github:cemheren/quicksheet-copilot-ext
copilot: test, 1, 1
```

![Copilot extension running on the desktop](copilot-ext-example.png)

You can reference cell ranges with `{A1::C10}` syntax and Copilot receives the actual cell contents as context:

```
copilot: summarize this data {B1::E50}, 3, 1
copilot: generate random test data with name and age, 2, 10
```

### Example: Weather forecast widget

The [quicksheet-weather](https://github.com/cemheren/quicksheet-weather) extension turns a cell into a live 7-day weather forecast:

```
ext: github:cemheren/quicksheet-weather
wthr: Seattle, 2, 7
```

### Build your own

Extensions are regular .NET (or any language) programs that read/write JSON lines on stdin/stdout. The protocol is intentionally minimal:

1. QuickSheet sends `{"type":"init"}` → extension replies with `{"type":"register", "prefix":"xyz", ...}`
2. When a user activates a cell matching the prefix, QuickSheet sends `{"type":"activate", ...}` → extension replies with `{"type":"write", "cells":[...]}` to fill the grid.

Ship a `quicksheet-extension.json` manifest in your repo and you're done. See the [weather extension](https://github.com/cemheren/quicksheet-weather) for a minimal working example, or the [Copilot extension](https://github.com/cemheren/quicksheet-copilot-ext) for something more advanced.

## Keyboard Shortcuts

### Navigation

| Shortcut | Action |
|----------|--------|
| Arrow Keys | Navigate cells |
| Shift+Arrow | Extend selection / multi-select |
| Tab | Move right, wrap to next row |
| Enter | Activate cell (open URL / run command) |

### Editing

| Shortcut | Action |
|----------|--------|
| F2 | Edit cell in-place |
| Backspace | Delete last character / clear selection |
| Delete | Clear cell or selection |
| Escape | Cancel edit |
| Type any character | Start editing the selected cell |

### View & Layout

| Shortcut | Action |
|----------|--------|
| F1 | Toggle raw vs resolved cell display |
| F3 | Rebuild grid / recalculate layout |
| F4 | Decrease column width |
| F5 | Increase column width |

### Clipboard & Search

| Shortcut | Action |
|----------|--------|
| Ctrl+C | Copy cell(s) |
| Ctrl+Shift+C | Copy resolved / displayed content |
| Ctrl+X | Cut cell(s) |
| Ctrl+V | Paste (supports multi-line) |
| Ctrl+F | Search |

### File & Row Operations

| Shortcut | Action |
|----------|--------|
| Ctrl+S | Save to CSV |
| Ctrl+D | Delete row |
| Ctrl+O | Insert row below |
| Ctrl+P | Insert row above |
| Ctrl+Q | Quit |

## License

MIT
