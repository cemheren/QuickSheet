# Extension Protocol Specification

> **Protocol version: 1** — QuickSheet v0.14.0+

This document is the strict reference for building QuickSheet extensions. Follow it exactly to avoid common pitfalls that break extensions on install.

---

## Table of Contents

1. [Overview](#overview)
2. [Transport Layer](#transport-layer)
3. [Lifecycle](#lifecycle)
4. [Messages: Host → Extension](#messages-host--extension)
5. [Messages: Extension → Host](#messages-extension--host)
6. [Cell Coordinate System](#cell-coordinate-system)
7. [Output Formats](#output-formats)
8. [Manifest File](#manifest-file)
9. [Environment Variables](#environment-variables)
10. [Rules & Constraints](#rules--constraints)
11. [Common Mistakes](#common-mistakes)
12. [Minimal Working Example](#minimal-working-example)

---

## Overview

Extensions are standalone programs (any language) that communicate with QuickSheet over **JSON-lines on stdin/stdout**. Each extension registers a cell prefix (e.g., `wthr`) and receives activation messages when the user types that prefix into a cell.

```
┌──────────────┐  stdin (JSON-lines)   ┌───────────────┐
│  QuickSheet  │ ────────────────────▶ │   Extension   │
│    (host)    │ ◀──────────────────── │   (process)   │
└──────────────┘  stdout (JSON-lines)  └───────────────┘
```

## Transport Layer

| Rule | Detail |
|------|--------|
| Format | One JSON object per line. No pretty-printing. No trailing commas. |
| Encoding | UTF-8, no BOM |
| Delimiter | `\n` (newline). Each message is exactly one line. |
| Direction | stdin = host-to-extension, stdout = extension-to-host |
| stderr | Logged by host for debugging. Never parsed as protocol messages. |
| Buffering | **Flush stdout after every message.** Buffered I/O will hang the protocol. |

**Critical:** Many languages buffer stdout by default when it's not a TTY (which it isn't — it's a pipe). You must explicitly flush after each `WriteLine`. In C# use `Console.Out.Flush()` or `AutoFlush = true`. In Python use `flush=True` or `sys.stdout.flush()`.

## Lifecycle

```
Host starts process
  │
  ├─▶ Host sends: {"type":"init","version":1}
  │
  ◀── Extension replies: {"type":"register","prefix":"xyz","name":"My Ext","version":"1.0.0"}
  │
  │   ... extension is now registered, prefix is live ...
  │
  ├─▶ Host sends: {"type":"activate","id":"activate-3-2","anchor":{"row":4,"col":2},"params":["arg1"],"gridCols":1,"gridRows":10}
  │
  ◀── Extension replies: {"type":"write","id":"activate-3-2","cells":[...]}
  │
  │   ... user deletes the cell or changes it ...
  │
  ├─▶ Host sends: {"type":"deactivate","id":"activate-3-2","anchor":{"row":4,"col":2}}
  │
  │   ... extension can clean up timers/state for that activation ...
  │
  │   ... host shuts down ...
  │
  ├─▶ Host closes stdin
  └── Extension should exit cleanly (host will kill after 3 seconds if not)
```

### Startup timing

- The host sends `init` immediately after launching the process.
- The extension **must** reply with `register` — not `init`, not `ready`. The type must be exactly `"register"`.
- The host will not send any `activate` messages until it receives the `register` response.
- If the extension does not register within a reasonable time (~5s), it may be killed.

## Messages: Host → Extension

### `init`

Sent once immediately after process start.

```json
{"type":"init","version":1}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"init"` | Message discriminator |
| `version` | `int` | Protocol version (currently `1`) |

### `activate`

Sent when a user writes a cell matching the extension's registered prefix.

```json
{
  "type": "activate",
  "id": "activate-3-2",
  "anchor": {"row": 4, "col": 2},
  "params": ["98112"],
  "gridCols": 2,
  "gridRows": 7
}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"activate"` | Message discriminator |
| `id` | `string` | Unique activation ID. **Echo this back** in your response. |
| `anchor` | `{row, col}` | The cell where output starts (one row below the command cell) |
| `params` | `string[]` | User-provided parameters (comma-separated in the cell, trimmed) |
| `gridCols` | `int` | How many columns wide the output area is (user-specified or default `1`) |
| `gridRows` | `int` | How many rows tall the output area is (user-specified or default `10`) |

**Important:** The `anchor` is already computed by the host. It is one row below the command cell. You do **not** need to compute offsets from the command cell yourself.

### `deactivate`

Sent when the command cell is deleted, cleared, or changed to something else.

```json
{
  "type": "deactivate",
  "id": "activate-3-2",
  "anchor": {"row": 4, "col": 2}
}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"deactivate"` | Message discriminator |
| `id` | `string` | The activation ID that is being deactivated |
| `anchor` | `{row, col}` | The anchor that was used |

Use this to stop timers, cancel HTTP requests, or free resources for that activation.

## Messages: Extension → Host

### `register`

**Must** be sent exactly once, as the first response to `init`.

```json
{"type":"register","prefix":"wthr","name":"Weather","version":"1.0.0"}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"register"` | Must be exactly this string |
| `prefix` | `string` | The cell prefix to claim (lowercase, no colon, no trailing space) |
| `name` | `string` | Human-readable extension name |
| `version` | `string` or `int` | Extension version |

**Rules for `prefix`:**
- Lowercase letters and numbers only. No spaces, no colons, no special characters.
- Must **not** include the trailing colon. Write `"wthr"` not `"wthr:"`.
- The host normalizes to lowercase. `"WTHR"` and `"wthr"` are the same prefix.
- Choose something short (3–6 chars) and unlikely to conflict.

### `write`

Sends cell values to display in the grid. This is the primary output mechanism.

```json
{"type":"write","id":"activate-3-2","cells":[["Mon 72°F ☀️"],["Tue 68°F 🌧️"]]}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"write"` | Message discriminator |
| `id` | `string` | **Must match** the `id` from the `activate` message |
| `cells` | array | Cell data in one of three formats (see [Output Formats](#output-formats)) |

**You may send multiple `write` messages** for the same activation ID. Each one overwrites the previous output. This enables live-updating extensions (timers, monitors, etc.).

### `status`

Displays a temporary status message at the anchor cell (e.g., "Loading...").

```json
{"type":"status","id":"activate-3-2","message":"Fetching forecast..."}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"status"` | Message discriminator |
| `id` | `string` | Must match the activation ID |
| `message` | `string` | Text to show (will be overwritten by next `write`) |

### `error`

Reports an error. The host displays `[err: message]` at the anchor cell and marks the cell as errored (red highlight).

```json
{"type":"error","id":"activate-3-2","message":"City not found"}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"error"` | Message discriminator |
| `id` | `string` | Must match the activation ID |
| `message` | `string` | Error description (keep it short — one cell width) |

### `log`

Debug logging. The host records it but does not display it to the user.

```json
{"type":"log","level":"info","message":"Cache hit for AAPL"}
```

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"log"` | Message discriminator |
| `level` | `string` | `"info"`, `"warn"`, or `"error"` |
| `message` | `string` | Debug text |

## Cell Coordinate System

All coordinates in `write` messages are **relative to the anchor** (0-based).

```
Command cell:  row=3, col=2  →  user types "wthr: 98112"
Anchor:        row=4, col=2  →  output starts here

Your write coords:
  (r=0, c=0) → grid cell (4, 2)   ← anchor
  (r=1, c=0) → grid cell (5, 2)
  (r=0, c=1) → grid cell (4, 3)
```

**Never use absolute grid coordinates.** The host adds your relative offsets to the anchor automatically. If you write to `r:0, c:0`, it appears at the anchor cell. If you write to `r:3, c:0`, it appears 3 rows below the anchor.

**Stay within bounds:** Only write within `gridRows` × `gridCols`. The host clips writes that exceed the grid boundaries, but well-behaved extensions should respect the declared area.

## Output Formats

The `cells` field in a `write` message accepts three formats:

### Format 1: Grid array (recommended for most extensions)

A 2D array where each inner array is a row of string values.

```json
{"cells": [["Mon 72°F", "☀️"], ["Tue 68°F", "🌧️"], ["Wed 75°F", "⛅"]]}
```

Maps to: row 0 = `["Mon 72°F", "☀️"]`, row 1 = `["Tue 68°F", "🌧️"]`, etc.
Each string becomes one cell at `(rowIndex, colIndex)` relative to anchor.

### Format 2: Flat string array (single-column output)

A flat array of strings. Each element becomes one row in column 0.

```json
{"cells": ["Line 1", "Line 2", "Line 3"]}
```

Maps to: `(0,0)="Line 1"`, `(1,0)="Line 2"`, `(2,0)="Line 3"`.

### Format 3: Object array (explicit positioning)

An array of `{r, c, v}` objects for sparse or non-contiguous output.

```json
{"cells": [{"r":0,"c":0,"v":"CPU"}, {"r":0,"c":1,"v":"45%"}, {"r":2,"c":0,"v":"RAM"}, {"r":2,"c":1,"v":"8.2 GB"}]}
```

**All three formats use 0-based relative coordinates.** The host maps them to absolute grid positions.

## Manifest File

Every extension repo must contain `quicksheet-extension.json` at the root:

```json
{
  "name": "weather",
  "version": "1.0.0",
  "prefix": "wthr",
  "description": "7-day weather forecast for a location",
  "entry": "dotnet run --project WeatherExt.csproj",
  "minProtocolVersion": 1
}
```

| Field | Required | Description |
|-------|----------|-------------|
| `name` | Yes | Extension identifier (alphanumeric + hyphens) |
| `version` | Yes | SemVer string |
| `prefix` | Yes | Same prefix that `register` will report. Must match exactly. |
| `description` | Yes | One-line description |
| `entry` | Yes | Shell command to start the extension process |
| `minProtocolVersion` | No | Minimum protocol version required (default: 1) |

**The `entry` command** is executed with the extension's repo directory as the working directory. It can be any command: `dotnet run`, `python main.py`, `node index.js`, `./binary`, etc.

## Environment Variables

The host sets these before launching the extension process:

| Variable | Value |
|----------|-------|
| `QUICKSHEET_EXTENSIONS_DIR` | Path to the extensions root directory |
| `QUICKSHEET_PROTOCOL_VERSION` | Protocol version as string (e.g., `"1"`) |

## Rules & Constraints

These are hard rules. Violating them causes broken behavior:

1. **Do not send data in a loop.** Extensions should respond to `activate` with one `write` (or a sequence of `status` → `write`). If you need live updates, use a timer with reasonable intervals (≥1 second). Never flood the host with writes — the UI repaints on each one.

2. **Always display output below the activation cell.** The host computes the anchor one row below the command cell. Your relative coordinates start at `(0, 0)` = anchor. Never write at negative row offsets (above the command cell).

3. **Respect `gridCols` and `gridRows`.** These define the output area the user allocated. Don't write beyond them. If the user specifies `gridCols=2, gridRows=5`, limit your output to 2 columns and 5 rows.

4. **Echo the activation `id` exactly.** Every `write`, `status`, and `error` message must include the `id` from the `activate` message. If the ID doesn't match, the host ignores the message.

5. **Flush stdout after every message.** Pipe buffering will make your extension appear dead.

6. **Reply to `init` with `register`, nothing else.** The first message you send must be `{"type":"register",...}`. Not `init`, not `ready`, not `hello`.

7. **Prefix must not include the colon.** Register `"wthr"`, not `"wthr:"`.

8. **Exit cleanly when stdin closes.** The host closes stdin on shutdown. Read until EOF, then exit. The host kills after 3 seconds if you don't.

9. **No NuGet/pip/npm dependencies at install time.** Extensions are `git clone`'d and run. If your entry command requires package restore, the user must have those packages pre-installed. Keep it minimal.

10. **One extension = one prefix.** Don't register multiple prefixes from a single process.

## Common Mistakes

| Mistake | Symptom | Fix |
|---------|---------|-----|
| Prefix includes colon (`"wthr:"`) | Extension registers but never activates | Remove the colon: `"wthr"` |
| Using absolute coordinates in cells | Output appears at wrong position or is invisible | Use 0-based relative coords. `(0,0)` = anchor cell. |
| Forgetting to flush stdout | Extension appears to hang after init | Add `Console.Out.Flush()` or set `AutoFlush = true` |
| Responding with `type: "response"` instead of `"write"` | No output appears | Use `"type": "write"` for cell output |
| Sending `type: "init"` back instead of `"register"` | Host never recognizes the extension | Reply with `"type": "register"` |
| Not echoing the `id` field | Write messages are silently dropped | Copy `id` from activate into your write/status/error |
| Writing above the anchor (negative row) | Overwrites the command cell or other data | Only use `r >= 0` in cell writes |
| Sending output before receiving `activate` | Output is lost (no active call to route it to) | Wait for `activate`, then respond |
| Flooding writes without delay | UI becomes unresponsive | Rate-limit to ≤1 write/second for live extensions |

## Minimal Working Example

A complete extension in C# (.NET 9, zero dependencies):

```csharp
using System.Text.Json;

// Disable stdout buffering
Console.OutputEncoding = System.Text.Encoding.UTF8;
var stdout = Console.Out;

string? line;
while ((line = Console.ReadLine()) != null)
{
    if (string.IsNullOrWhiteSpace(line)) continue;

    using var doc = JsonDocument.Parse(line);
    var root = doc.RootElement;
    string type = root.GetProperty("type").GetString() ?? "";

    switch (type)
    {
        case "init":
            stdout.WriteLine(JsonSerializer.Serialize(new
            {
                type = "register",
                prefix = "hello",
                name = "Hello World",
                version = "1.0.0"
            }));
            stdout.Flush();
            break;

        case "activate":
            string id = root.GetProperty("id").GetString() ?? "";
            string[] parms = root.GetProperty("params")
                .EnumerateArray()
                .Select(p => p.GetString() ?? "")
                .ToArray();

            string greeting = parms.Length > 0
                ? $"Hello, {parms[0]}!"
                : "Hello, World!";

            stdout.WriteLine(JsonSerializer.Serialize(new
            {
                type = "write",
                id,
                cells = new[] { new[] { greeting } }
            }));
            stdout.Flush();
            break;

        case "deactivate":
            // Nothing to clean up
            break;
    }
}
```

**Manifest** (`quicksheet-extension.json`):
```json
{
  "name": "hello",
  "version": "1.0.0",
  "prefix": "hello",
  "description": "Greets the user",
  "entry": "dotnet run",
  "minProtocolVersion": 1
}
```

**Usage:** Type `hello: World` in any cell → output: `Hello, World!` one row below.

### Python equivalent

```python
import sys
import json

for line in sys.stdin:
    line = line.strip()
    if not line:
        continue
    msg = json.loads(line)

    if msg["type"] == "init":
        print(json.dumps({
            "type": "register",
            "prefix": "hello",
            "name": "Hello World",
            "version": "1.0.0"
        }), flush=True)

    elif msg["type"] == "activate":
        params = msg.get("params", [])
        greeting = f"Hello, {params[0]}!" if params else "Hello, World!"
        print(json.dumps({
            "type": "write",
            "id": msg["id"],
            "cells": [[greeting]]
        }), flush=True)
```

---

## Protocol Version History

| Version | Introduced | Changes |
|---------|-----------|---------|
| 1 | v0.14.0 | Initial protocol. `init`, `activate`, `deactivate`, `register`, `write`, `status`, `error`, `log`. |

---

*Last updated: 2026-05-16. Source of truth: `Extensions/ExtensionProtocol.cs` and `Extensions/ExtensionManager.cs` in the QuickSheet repo.*
