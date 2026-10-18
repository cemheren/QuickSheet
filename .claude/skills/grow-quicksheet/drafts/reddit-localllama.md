# Reddit r/LocalLLaMA + r/MachineLearning draft — QuickSheet as AI-cost dashboard

Drafted 2026-06-14. These subs reward practical tools that solve a real annoyance.
Frame: "I was losing track of API spend across 4 different AI CLIs, so I turned my
wallpaper into a live cost monitor."

---

## r/LocalLLaMA

### Title

**I turned my desktop wallpaper into a live AI-cost dashboard — tracks Claude, Copilot, Ollama, and Aider token usage without a browser tab**

### Body

```
I run Claude Code, GitHub Copilot CLI, Aider, and a local Ollama instance
depending on the task. The problem: I kept losing track of what I'd spent on
API calls each day. Token dashboards live in browser tabs I never look at.

So I built a spreadsheet that IS my desktop wallpaper. It's always visible
behind every window — no tab-switching, no tray icon, just the data sitting
there.

Here's what the AI-workflow layout looks like:

- Row of `r:` cells that launch each tool with one keystroke (Claude, Copilot
  suggest/explain, Aider --ask)
- An `i:` cell that runs `claude /cost` and streams token counts directly into
  the wallpaper (updates every time I glance at it)
- A sparkline row showing daily spend: `s: 12,18,22,15,9,7,14` renders as
  ▂▃▆▁▅█ — one week at a glance
- An Ollama extension cell (`ext: github:cemheren/quicksheet-ollama`) that
  shows which models are loaded and their VRAM usage

The "live subprocess" feature (`i: <command>`) is what makes it work as a
monitor — it runs the command and pipes stdout back into the cell, live. ConPTY
on Windows, pipe redirect on Linux. Caps at 200 lines so it doesn't eat RAM.

For the local LLM crowd specifically:
- `i: ollama ps` → shows loaded models + VRAM in the wallpaper
- `i: nvidia-smi --query-gpu=utilization.gpu,memory.used --format=csv,noheader`
  → GPU load always visible
- Pair with a `L:` (loop) prefix to auto-refresh every 30s

It's a .NET 9 app. Zero external dependencies — all native interop is
hand-written P/Invoke against X11/WinForms. Works on Linux (X11) and Windows.

dotnet run -- examples/ai-workflow.csv --desktop

GitHub: https://github.com/cemheren/QuickSheet

Would love feedback from people running local models — what other metrics would
you want pinned to the wallpaper? I'm thinking `i: litellm --cost` or a cell
that tails the Ollama server log for request counts.
```

### First comment (post immediately after submitting)

```
A few details for the "how does this actually work" crowd:

1. The wallpaper embedding uses _NET_WM_WINDOW_TYPE_DESKTOP on X11 (sets the
   window to the desktop layer). On Windows it finds the WorkerW handle behind
   the icons and parents itself to it.

2. The `i:` (inline process) cells spawn a real subprocess. Output is captured
   line-by-line and rendered in the cell. It's not polling an API — it's
   literally `Process.Start` with stdout redirected.

3. Extensions use a JSON-lines protocol over stdin/stdout. The Ollama extension
   is just a .NET console app that calls localhost:11434/api/tags and formats
   the response.

4. The whole thing autosaves to CSV every 5 seconds. Your layout persists
   across reboots.

No Electron, no web tech, no NuGet packages. Just P/Invoke and spite.
```

---

## r/MachineLearning [D] Discussion

### Title

**[D] QuickSheet — open-source ambient AI-cost monitor that runs as your desktop wallpaper**

### Body

```
Not a paper, just a practical tool I've been using daily:

The problem it solves: when you're running experiments with multiple AI APIs
(Claude, GPT, local Ollama), it's easy to lose track of daily/weekly spend.
Dashboard tabs get buried. QuickSheet puts the numbers directly on your desktop
wallpaper where they're always visible.

Key ML-relevant features:
- `i: curl -s localhost:11434/api/tags | jq '.models[].name'` → live list of
  loaded Ollama models, always visible
- `i: nvidia-smi ...` → GPU utilization in a cell, auto-refreshing
- Sparkline cells for tracking daily token spend visually
- Extension protocol for custom data sources (wrote one for Ollama model
  status in ~50 lines)

It's a .NET 9 desktop app that embeds as the wallpaper layer (X11 on Linux,
WorkerW on Windows). Zero dependencies, builds from source in one command.

Useful for anyone who:
- Runs local models and wants GPU/VRAM stats always visible
- Uses multiple AI CLIs and wants a unified cost glance
- Wants runnable command cells (click to launch training scripts, etc.)

https://github.com/cemheren/QuickSheet

Happy to answer questions about the implementation — the inline subprocess
system and extension protocol might be interesting to ML-tooling people.
```

---

## Posting notes

- **r/LocalLLaMA**: Best time is weekday morning US (9-11am ET). This sub
  loves "here's my workflow" posts with practical tools. Lead with the problem
  (cost tracking), not the project. Flair: "Resources" or "Other".
- **r/MachineLearning**: Use [D] tag. Keep it concise and technical. This crowd
  is allergic to hype — the "No Electron, no web tech" angle plays well here.
- **Screenshot required**: Need the AI-workflow wallpaper screenshot to post.
  Without it, the post has no hook. GATED on asset.
- **Cross-post timing**: Post to r/LocalLLaMA first (more receptive). If it
  gets traction (>20 upvotes in 2h), post r/MachineLearning the next day.
- **Don't also post to r/programming or r/commandline the same week** — Reddit
  sees cross-posting and some users will call it out.
