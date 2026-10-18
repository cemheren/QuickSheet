# Wayland Support Design Document

> Investigation for [#3](https://github.com/cemheren/QuickSheet/issues/3) — Wayland support for `--desktop` mode.

## Summary

QuickSheet's Linux desktop mode currently uses raw X11 P/Invoke (`libX11.so.6`, `libXft.so.2`) with the `_NET_WM_WINDOW_TYPE_DESKTOP` hint. Under Wayland, this runs through XWayland but the desktop-type hint is not honored — the window renders as a normal floating window.

Real Wayland support is achievable via the **wlr-layer-shell** protocol using hand-written P/Invoke against `libwayland-client.so.0` (a standard system library). This document scopes the implementation.

## Compositor Support Matrix

| Compositor | wlr-layer-shell | Notes |
|---|---|---|
| **Sway** | ✅ Full | wlroots-native, primary target |
| **Hyprland** | ✅ Full | wlroots-based, large r/unixporn audience |
| **KDE Plasma (KWin)** | ✅ Full | Supported since Plasma 5.x; default session is Wayland in Plasma 6+ |
| **Labwc, Wayfire, River** | ✅ Full | wlroots-based compositors |
| **Weston** | ✅ | Reference compositor |
| **GNOME (Mutter)** | ❌ None | No plans to implement; uses private shell protocol |

**Coverage estimate:** wlr-layer-shell reaches ~70–80% of Wayland desktop users (everyone except GNOME). The X11/XWayland fallback remains for GNOME users.

## Protocol Overview

The [`zwlr_layer_shell_v1`](https://wayland.app/protocols/wlr-layer-shell-unstable-v1) protocol (version 4) defines:

### Interfaces

**`zwlr_layer_shell_v1`** — factory for layer surfaces:
- `get_layer_surface(surface, output, layer, namespace)` — creates a layer surface
- `destroy()` — cleanup (since v3)

**`zwlr_layer_surface_v1`** — the positioned surface:
- `set_size(width, height)` — 0,0 means "fill anchored dimension"
- `set_anchor(flags)` — bitfield: top|bottom|left|right (all 4 = fullscreen)
- `set_exclusive_zone(zone)` — -1 for backgrounds (don't reserve space)
- `set_margin(top, right, bottom, left)`
- `set_keyboard_interactivity(mode)` — on_demand (v4) for our input needs
- `ack_configure(serial)` — required before first buffer attach
- `destroy()`

**Events:**
- `configure(serial, width, height)` — compositor tells us the surface size
- `closed()` — compositor wants us gone

### Layer Enum

```
background = 0   ← QuickSheet target (below all windows)
bottom     = 1
top        = 2
overlay    = 3
```

## Proposed Architecture

```
Platform/Linux/
├── DesktopWindow.cs         (existing X11 implementation — unchanged)
├── LinuxDesktopHost.cs      (add runtime detection: X11 vs Wayland)
├── X11Methods.cs            (existing — unchanged)
├── Wayland/
│   ├── WaylandMethods.cs    (P/Invoke for libwayland-client.so.0)
│   ├── LayerShellMethods.cs (protocol opcodes for zwlr_layer_shell)
│   ├── WaylandWindow.cs     (equivalent of DesktopWindow for Wayland)
│   └── ShmBuffer.cs         (wl_shm shared-memory buffer management)
```

### Runtime Detection

```csharp
// In LinuxDesktopHost.cs
bool useWayland = Environment.GetEnvironmentVariable("WAYLAND_DISPLAY") != null
    && !Environment.GetEnvironmentVariable("QUICKSHEET_FORCE_X11")?.Equals("1") == true;
```

If `WAYLAND_DISPLAY` is set AND `libwayland-client.so.0` loads AND the compositor advertises `zwlr_layer_shell_v1` in its registry → use Wayland path. Otherwise fall back to X11 (current behavior).

### Key P/Invoke Surface

```csharp
internal static class WaylandMethods
{
    private const string Lib = "libwayland-client.so.0";

    [DllImport(Lib)] public static extern IntPtr wl_display_connect(string? name);
    [DllImport(Lib)] public static extern void wl_display_disconnect(IntPtr display);
    [DllImport(Lib)] public static extern int wl_display_dispatch(IntPtr display);
    [DllImport(Lib)] public static extern int wl_display_roundtrip(IntPtr display);
    [DllImport(Lib)] public static extern IntPtr wl_display_get_registry(IntPtr display);
    [DllImport(Lib)] public static extern int wl_display_flush(IntPtr display);

    // Registry
    [DllImport(Lib)] public static extern IntPtr wl_registry_bind(
        IntPtr registry, uint name, ref WlInterface iface, uint version);

    // Proxy (for marshalling protocol requests)
    [DllImport(Lib)] public static extern IntPtr wl_proxy_marshal_constructor(
        IntPtr proxy, uint opcode, ref WlInterface iface, __arglist);
    [DllImport(Lib)] public static extern void wl_proxy_marshal(
        IntPtr proxy, uint opcode, __arglist);
    [DllImport(Lib)] public static extern int wl_proxy_add_listener(
        IntPtr proxy, IntPtr listeners, IntPtr data);
    [DllImport(Lib)] public static extern void wl_proxy_destroy(IntPtr proxy);
}
```

### Rendering Strategy

Two options (in preference order):

1. **SHM buffers + CPU rendering** — Same approach as current X11 (XDrawRectangle/XftDrawString equivalent). Allocate a `wl_shm_pool`, create `wl_buffer`, write ARGB pixels directly. This matches the zero-dep constraint perfectly.

2. **EGL/OpenGL** — Overkill for text grid rendering and would require `libEGL.so`/`libGLESv2.so` P/Invoke. Not recommended.

**Font rendering:** Without Xft, we need an alternative for anti-aliased text. Options:
- P/Invoke `libfontconfig.so.1` + `libfreetype.so.6` (available on all graphical Linux) to rasterize glyphs into the SHM buffer.
- Use a bitmap/monospace approach (render glyphs once, blit). Simpler and faster for a character grid.

### Input Handling

Layer shell v4 adds `keyboard_interactivity: on_demand` — the surface can receive keyboard focus when clicked. This is essential for QuickSheet's editing.

Mouse events come through the standard `wl_pointer` interface (part of `wl_seat`).

## Implementation Phases

### Phase 1: Proof of Concept (smallest viable change)
- Add `WaylandMethods.cs` with basic display/registry P/Invoke.
- Detect Wayland at runtime; print "Wayland detected, layer-shell available" or fall back.
- Create a colored background surface using wlr-layer-shell (no text, just solid color).
- **Goal:** Confirm the P/Invoke approach works without wayland-scanner.

### Phase 2: Grid Rendering
- Port the grid layout/sizing logic from `DesktopWindow.cs` into `WaylandWindow.cs`.
- Implement SHM buffer allocation and pixel-level drawing.
- Render cell borders, selection highlight, text (FreeType rasterization).

### Phase 3: Input & Feature Parity
- Implement keyboard/mouse input via `wl_seat`/`wl_keyboard`/`wl_pointer`.
- Port editing, search, extensions, inline processes.
- Handle `configure` events for output size changes / multi-monitor.

### Phase 4: Polish
- Handle compositor `closed` event gracefully.
- Support `QUICKSHEET_FORCE_X11=1` override.
- Add `--wayland` / `--x11` explicit flags.
- Update README and help text.

## Constraints & Decisions

| Constraint | Decision |
|---|---|
| Zero NuGet deps | All interop via hand-written P/Invoke against system libs |
| No regression to X11 | Wayland path is opt-in via runtime detection |
| GNOME unsupported | Document limitation; X11/XWayland fallback works |
| Font rendering without Xft | FreeType + Fontconfig P/Invoke (both are universal system libs) |
| Build must stay cross-platform | New files under `Platform/Linux/Wayland/`, guarded by `PLATFORM_LINUX` |

## Open Questions

1. **Should Phase 1 land as a single PR or be split further?** Runtime detection alone (no surface creation) could be a standalone commit.
2. **FreeType vs bitmap font:** Bitmap is simpler but less flexible. FreeType gives parity with Xft quality. Recommend FreeType since the grid is fixed-pitch anyway.
3. **Multi-monitor:** wlr-layer-shell can target a specific `wl_output`. Should QuickSheet span all outputs or pick one? Current X11 impl uses full root window (spans all). Recommend: default to primary output, add `--output` flag later.

## References

- [wlr-layer-shell protocol spec](https://wayland.app/protocols/wlr-layer-shell-unstable-v1)
- [Protocol XML source](https://github.com/swaywm/wlr-protocols/blob/master/unstable/wlr-layer-shell-unstable-v1.xml)
- [wlroots layer-shell C example](https://github.com/swaywm/wlroots/blob/master/examples/layer-shell.c)
- [Wayland client P/Invoke pattern](https://wayland.freedesktop.org/docs/html/ch05.html) — wire protocol basics
- Current X11 implementation: `Platform/Linux/DesktopWindow.cs`, `Platform/Linux/X11Methods.cs`
