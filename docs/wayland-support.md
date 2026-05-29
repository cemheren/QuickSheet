# Wayland Support Investigation

> Tracks: [#3](https://github.com/cemheren/QuickSheet/issues/3)

## Summary

QuickSheet's Linux desktop mode currently requires X11 (`_NET_WM_WINDOW_TYPE_DESKTOP`
hint via raw `libX11.so.6` P/Invoke). Under Wayland, this hint is not honored —
the window renders as a normal floating window through XWayland.

Real Wayland support is **feasible** for wlroots-based compositors (Sway, Hyprland,
Wayfire) via the `wlr-layer-shell-unstable-v1` protocol. GNOME/Mutter does **not**
implement layer-shell and has no public wallpaper-surface protocol, making it
unsupported without a workaround.

## Compositor Compatibility Matrix

| Compositor      | Protocol Available             | Feasibility |
|-----------------|-------------------------------|-------------|
| Sway            | `wlr-layer-shell` (background layer) | ✅ Supported |
| Hyprland        | `wlr-layer-shell` (background layer) | ✅ Supported |
| Wayfire         | `wlr-layer-shell` (background layer) | ✅ Supported |
| KDE Plasma      | `org_kde_plasma_window_management`    | ⚠️ Possible (different protocol) |
| GNOME (Mutter)  | None — no layer-shell, no public API  | ❌ Not feasible without extension |

## Implementation Approach

### Runtime Detection

```
XDG_SESSION_TYPE=wayland  →  try Wayland path
WAYLAND_DISPLAY set       →  try wl_display_connect()
Fallback                  →  existing X11 path (unchanged)
```

In `Program.cs`, the existing Wayland detection already checks `XDG_SESSION_TYPE`.
Instead of just warning and falling through to XWayland, a Wayland host could be
attempted first.

### Required P/Invoke Surface (zero NuGet deps)

All functions live in `libwayland-client.so.0` (present on every Wayland system):

```csharp
[DllImport("libwayland-client.so.0")]
static extern IntPtr wl_display_connect(string? name);

[DllImport("libwayland-client.so.0")]
static extern void wl_display_disconnect(IntPtr display);

[DllImport("libwayland-client.so.0")]
static extern IntPtr wl_display_get_registry(IntPtr display);

[DllImport("libwayland-client.so.0")]
static extern int wl_display_roundtrip(IntPtr display);

[DllImport("libwayland-client.so.0")]
static extern int wl_display_dispatch(IntPtr display);

[DllImport("libwayland-client.so.0")]
static extern IntPtr wl_registry_bind(IntPtr registry, uint name,
    IntPtr iface, uint version);
```

### The Layer-Shell Problem

`wlr-layer-shell` is **not** exposed as a standalone `.so` — it's a Wayland
protocol extension compiled into compositors. Clients need the protocol's
`wl_interface` struct pointer to bind via `wl_registry_bind`.

**Options to get the interface pointer:**

1. **Thin C shim** — compile a ~20-line `.so` that exports the generated
   `zwlr_layer_shell_v1_interface` symbol. Ship as `libqs-layer-shell.so`.
   This is how `swaybg` and `waybar` work internally.

2. **Reconstruct in C#** — manually define the `wl_interface` struct layout
   (name, version, method count, method signatures). Fragile but avoids an
   external build step.

3. **dlopen + dlsym** — if the compositor exposes the symbol in a shared lib,
   load it at runtime. Unreliable across distros.

**Recommendation:** Option 1 (thin C shim). It's ~50 lines of C, builds with a
single `cc` invocation, and is stable across compositor versions. The shim would
be checked into `Platform/Linux.Wayland/native/` with a Makefile.

### Architecture

```
Platform/
├── Linux/              ← existing X11 path (unchanged)
│   ├── LinuxDesktopHost.cs
│   ├── DesktopWindow.cs
│   └── X11Methods.cs
└── Linux.Wayland/      ← new Wayland path
    ├── WaylandDesktopHost.cs   (implements IDesktopHost)
    ├── WaylandMethods.cs       (P/Invoke declarations)
    ├── LayerShellSurface.cs    (layer-shell surface management)
    └── native/
        ├── layer-shell-shim.c  (thin C shim for wl_interface)
        └── Makefile
```

### Rendering

The X11 path uses `libXft.so.2` for text rendering directly to an X11 window.
The Wayland equivalent would need:

- A shared-memory buffer (`wl_shm`) for the surface pixels
- Software text rendering (e.g., `libfreetype.so.6` + `libfontconfig.so.1`
  via P/Invoke, or a minimal bitmap font renderer)
- Double-buffering with `wl_surface_attach` / `wl_surface_commit`

This is significantly more work than X11 (where Xft handles font rendering),
but still achievable with zero NuGet deps since FreeType is a system library.

## Effort Estimate

| Component                     | Effort  | Risk   |
|-------------------------------|---------|--------|
| Runtime detection + fallback  | Small   | Low    |
| Wayland connection + registry | Medium  | Low    |
| Layer-shell binding (shim)    | Medium  | Medium |
| SHM buffer management         | Medium  | Low    |
| Text rendering (FreeType)     | Large   | Medium |
| Input handling (wl_keyboard)  | Medium  | Low    |
| **Total**                     | **~2-3 weeks** | **Medium** |

## Recommended Phased Plan

### Phase 1: Detection + graceful behavior (small PR)
- Detect Wayland session reliably
- If `wl_display_connect` succeeds and layer-shell is advertised, print
  "Wayland layer-shell support coming soon" instead of the current warning
- No functional change yet

### Phase 2: Layer-shell surface (core)
- Thin C shim for protocol binding
- Create background-layer surface, full-screen, below everything
- Render a solid color + basic text as proof-of-concept
- Target: Sway first (easiest to test, well-documented)

### Phase 3: Full rendering parity
- FreeType P/Invoke for text rendering into SHM buffers
- Port grid drawing logic from `DesktopWindow.cs`
- Input handling via `wl_keyboard` / `wl_pointer`
- Autosave, extensions, all features

### Phase 4: KDE Plasma (stretch)
- Investigate `org_kde_plasma_window_management` for Plasma users
- Likely a separate shim + protocol binding

## GNOME Workaround (Low Priority)

For GNOME users, the only viable path is:
- Render the grid to a PNG periodically
- Set it as wallpaper via `gsettings set org.gnome.desktop.background picture-uri`
- No interactivity (click/keyboard) — read-only dashboard mode
- Could be offered as a `--gnome-wallpaper` flag

## References

- [wlr-layer-shell protocol XML](https://github.com/swaywm/wlr-protocols/blob/master/unstable/wlr-layer-shell-unstable-v1.xml)
- [swaybg source](https://github.com/swaywm/swaybg) — reference C implementation
- [WaylandSharp](https://github.com/Drakulix/WaylandSharp) — outdated C# bindings (reference only)
- [wl_shm documentation](https://wayland-book.com/surfaces/shared-memory.html)
- [FreeType P/Invoke examples](https://github.com/nicbarker/clay) — general pattern

## Testing Requirements

A real Wayland session (Sway or Hyprland recommended) is needed to test.
The maintainer runs X11, so a contributor with a Wayland desktop would be
ideal for Phase 2+.
