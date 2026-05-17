# awesome-list fit analysis (2026-05-17)

## TL;DR (5 lines)

- Backlog item "submit to **awesome-selfhosted**" is a bad fit. That list is for *server* software (Plex, Nextcloud, Pi-hole, Gitea). QuickSheet is a desktop wallpaper app — does not pass their "must be a self-hosted service" bar. **Drop it.**
- Already-drafted (in `drafts/awesome-lists.md`): awesome-csharp, awesome-dotnet, awesome-tuis, awesome-cli-apps, terminaltrove, console.dev. Three of those (csharp, dotnet, tuis) are real fits; the others are *terminal*-focused and we now lead with the wallpaper angle.
- Best **new** lists to target, ranked: **Awesome-Linux-Software** (productivity → Office; QuickSheet --desktop on X11 fits), **Awesome-Windows** (productivity/utilities; WorkerW desktop trick fits), **awesome-csv** (QuickSheet is a CSV editor with extras), **awesome-i3** / **awesome-hyprland** / **awesome-river** (window-manager-adjacent rice lists where wallpaper tooling fits).
- DO NOT submit to: awesome-selfhosted, awesome-dashboards (browser-focused), awesome-rainmeter (essentially dormant, project-specific to Rainmeter skins).
- Net action: drop awesome-selfhosted from backlog; add Awesome-Linux-Software draft as the first new draft (highest unique-fit value).

## Per-list assessment

### Lists already drafted in `drafts/awesome-lists.md`

| List | Fit | Notes |
|------|-----|-------|
| awesome-csharp                            | ✅ strong | language-based; project is C#; trivial fit |
| awesome-dotnet                             | ✅ strong | same |
| rothgar/awesome-tuis                       | ⚠ marginal | borderline now that we lead with wallpaper mode; submit but lead with "TUI fallback when no --desktop" framing, not the desktop angle |
| awesome-cli-apps                           | ⚠ weak     | QuickSheet has no CLI; only `--export-md` is CLI-shaped |
| terminaltrove (directory)                  | ⚠ marginal | same TUI tension as awesome-tuis |
| console.dev (newsletter)                   | ⚠ marginal | same |

### Lists to ADD (not yet drafted)

| List | Fit | Why |
|------|-----|-----|
| **luong-komorebi/Awesome-Linux-Software** | ✅ strong | Productivity / Office sections; QuickSheet --desktop on X11 is exactly the kind of desktop-app this list catalogues. ~22k stars. Active. |
| **0PandaDEV/awesome-windows** (~2.4k★, active) | ✅ medium | Productivity / Customization sections; QuickSheet WorkerW desktop integration is the angle. Active list. Maintainer is explicit about quality bar — keep the entry sober. Older `Awesome-Windows/Awesome` link is dead (404). |
| **secretGeek/awesome-csv** (or similar)    | ✅ strong | QuickSheet IS a CSV viewer/editor with grid; fits the "CSV tools" niche directly. |
| **iggredible/Awesome-Vim** / vim-spreadsheet adjacency | ❌ | not vim |
| **awesome-hyprland** / **awesome-i3**     | ⚠ marginal | rice-aesthetic adjacency; submit only if a wallpaper screenshot exists |
| **awesome-rainmeter**                     | ❌ | dormant; specific to Rainmeter skins, not standalone apps |
| **awesome-dashboards**                    | ❌ | browser-focused (Grafana etc.) |
| **awesome-selfhosted**                    | ❌ | server software only; QuickSheet is a desktop app — would be rejected |

## Implications

1. **Drop `awesome-selfhosted` PR draft from queue.** Will be rejected; wastes a contribution slot in the list's review queue and signals not-understanding-the-list.
2. **Write a new draft (this run): `drafts/awesome-linux-software.md`** — Awesome-Linux-Software submission. Productivity section. Lead with `--desktop` wallpaper mode, mention CSV/extensions, MIT, zero deps. Most novel of the additions and highest hit-prob.
3. **Future run: draft `awesome-windows.md`** — Awesome-Windows submission. Same pitch with WorkerW emphasis.
4. **Future run: identify the canonical `awesome-csv` list** (a few exist) and draft a row for the best-maintained one.
5. **Skip the rice-list submissions** until a real wallpaper screenshot exists; rice lists are screenshot-driven and a text-only PR would land flat.
