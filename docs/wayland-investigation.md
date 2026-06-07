# Wayland Support Investigation

> Tracking issue: [#3](https://github.com/cemheren/QuickSheet/issues/3)

## Current State

QuickSheet's Linux desktop mode uses raw X11 P/Invoke (`libX11.so.6`, `libXft.so.2`)
and sets `_NET_WM_WINDOW_TYPE_DESKTOP` to embed beneath desktop icons. Under Wayland,
the app runs through XWayland but the window-type hint is **not honored** by most
compositors—the window renders as a regular floating surface.

## Wayland Wallpaper Landscape

There is **no standard Wayland protocol** for setting or embedding desktop backgrounds.
Each compositor family handles it differently:

| Compositor family | Protocol / mechanism | Interactive input? |
|---|---|---|
| wlroots (Sway, Wayfire, River, Hyprland) | `wlr-layer-shell-unstable-v1` — `BACKGROUND` layer | ✅ Yes (if `keyboard_interactivity` set) |
| KDE Plasma (KWin) | Internal Plasma Shell — no public protocol for third-party wallpaper apps | ❌ No API |
| GNOME (Mutter) | Internal GNOME Shell — no public protocol | ❌ No API |
| Cosmic (System76) | `wlr-layer-shell` supported (wlroots-based) | ✅ Yes |

**Key insight:** Only wlroots-based compositors expose a usable protocol
(`wlr-layer-shell`) for third-party apps to draw on the desktop background layer
*with* keyboard/mouse input. KDE and GNOME have no mechanism for this.

## The `wlr-layer-shell` Approach

### How it works

1. Connect to Wayland display (`wl_display_connect`)
2. Bind to `zwlr_layer_shell_v1` from the registry
3. Create a `wl_surface` → get a `zwlr_layer_surface_v1` on `LAYER_BACKGROUND`
4. Anchor to all four edges (fullscreen)
5. Set `keyboard_interactivity` to receive key events
6. Allocate shared-memory buffers, render grid, commit
7. Run `wl_display_dispatch` event loop

### What this enables

- Full-screen background surface beneath all windows ✅
- Keyboard input for cell editing ✅
- Mouse/pointer input for cell selection ✅
- Multi-monitor (one surface per `wl_output`) ✅
- Transparency (compositor-dependent, usually works) ✅

### Reference implementations

- **swaybg** — static wallpaper using layer-shell background layer
- **mpvpaper** — video wallpaper with layer-shell
- **hyprpaper** — multi-output wallpaper for Hyprland

## Implementation Options for QuickSheet

### Option A: Native Wayland P/Invoke (recommended)

Mirror the existing X11 approach but target `libwayland-client.so`:

```
Platform/Linux/
├── LinuxDesktopHost.cs        (detect X11 vs Wayland at runtime)
├── X11Methods.cs              (existing)
├── DesktopWindow.cs           (existing X11 window)
├── WaylandMethods.cs          (NEW: P/Invoke for libwayland-client + layer-shell)
└── WaylandDesktopWindow.cs    (NEW: layer-shell background surface)
```

**P/Invoke surface area** (minimum viable):

| Library | Functions needed |
|---|---|
| `libwayland-client.so` | `wl_display_connect`, `wl_display_disconnect`, `wl_display_dispatch`, `wl_display_roundtrip`, `wl_display_get_registry`, `wl_registry_bind`, `wl_compositor_create_surface`, `wl_surface_commit`, `wl_surface_attach`, `wl_surface_damage` |
| Generated layer-shell stubs | `zwlr_layer_shell_v1_get_layer_surface`, `zwlr_layer_surface_v1_set_size`, `zwlr_layer_surface_v1_set_anchor`, `zwlr_layer_surface_v1_set_keyboard_interactivity`, `zwlr_layer_surface_v1_ack_configure` |
| `libwayland-client.so` (SHM) | `wl_shm_create_pool`, `wl_shm_pool_create_buffer` — for pixel buffer rendering |

**Complexity estimate:** ~800–1200 lines of new C# (P/Invoke declarations + event
loop + SHM buffer management + text rendering via raw pixel writes or Pango/FreeType).

**Biggest challenge:** Text rendering. X11 path uses `libXft` for font rendering.
Under Wayland, there's no equivalent "just draw text" API—you must either:
- P/Invoke into FreeType2 + render glyphs to the SHM buffer directly
- P/Invoke into Pango/Cairo for full text layout
- Use a pre-rendered bitmap font (simplest, but ugly)

### Option B: C shim library

Write a small C helper (`libqs-wayland.so`) that wraps the layer-shell setup and
exposes a simple API:

```c
void* qs_wayland_init(int width, int height);
void  qs_wayland_draw(void* ctx, uint32_t* pixels, int w, int h);
int   qs_wayland_poll_event(void* ctx, struct qs_event* out);
void  qs_wayland_destroy(void* ctx);
```

QuickSheet would P/Invoke into this shim. This avoids the complexity of marshaling
Wayland's callback-heavy protocol in C#.

**Pros:** Simpler C# code, easier to maintain.  
**Cons:** Adds a build dependency (must compile C code), distributing .so file.

### Option C: XWayland fallback improvements

Instead of full native Wayland support, improve the XWayland experience:
- Detect XWayland and attempt `_NET_WM_STATE_BELOW` + skip-taskbar hints
- Auto-resize when output changes
- Document the limitations clearly

**Pros:** Zero new code for Wayland itself.  
**Cons:** Not a real solution—window still floats above actual wallpaper.

## Runtime Detection

```csharp
bool IsWayland() =>
    Environment.GetEnvironmentVariable("WAYLAND_DISPLAY") != null;

bool HasLayerShell() =>
    // After connecting, check wl_registry for zwlr_layer_shell_v1
    ...;
```

Strategy: Try Wayland + layer-shell first. Fall back to X11 if layer-shell is
unavailable (KDE/GNOME on Wayland). Fall back to XWayland warning if neither works.

## Compositor Support Matrix

| Compositor | Layer-shell? | QuickSheet would work? |
|---|---|---|
| Sway | ✅ | ✅ Full support |
| Wayfire | ✅ | ✅ Full support |
| River | ✅ | ✅ Full support |
| Hyprland | ✅ | ✅ Full support |
| Cosmic | ✅ | ✅ Full support |
| KDE Plasma (Wayland) | ❌ | ❌ No API — falls back to X11/XWayland |
| GNOME (Wayland) | ❌ | ❌ No API — falls back to X11/XWayland |
| Weston | ❌ | ❌ Reference compositor, no layer-shell |

## Recommended Roadmap

1. **Phase 1 (low effort):** Add runtime Wayland detection + clear user-facing
   message: "Wayland detected. Desktop mode requires a wlr-layer-shell compatible
   compositor (Sway, Hyprland, Wayfire, River). Falling back to XWayland."

2. **Phase 2 (medium effort):** Implement `WaylandMethods.cs` P/Invoke bindings
   for `libwayland-client.so`. Create background layer surface. Use a simple
   bitmap font or FreeType for initial text rendering.

3. **Phase 3 (full feature parity):** Multi-monitor support, proper font rendering
   (FreeType2/HarfBuzz), input handling for cell editing, mouse selection.

## Open Questions

- **Font rendering path:** FreeType2 P/Invoke vs Cairo/Pango vs bitmap font?
- **Keyboard layout handling:** Wayland uses `xkbcommon` for keymap processing—
  need P/Invoke into `libxkbcommon.so` for correct key translation.
- **Clipboard:** Wayland clipboard is async and protocol-based (`wl_data_device`).
  Needs separate implementation from X11's `XGetSelectionOwner`/`XConvertSelection`.
- **Distribution:** Should the layer-shell protocol XML be bundled and generated
  at build time, or should pre-generated C# stubs be committed?

## References

- [wlr-layer-shell protocol spec](https://wayland.app/protocols/wlr-layer-shell-unstable-v1)
- [swaybg source](https://github.com/swaywm/swaybg) — reference background client
- [mpvpaper](https://github.com/GhostNaN/mpvpaper) — video wallpaper via layer-shell
- [WaylandSharp](https://github.com/Drakulix/WaylandSharp) — archived C# Wayland bindings (reference only)
- [wlroots.net](https://github.com/varmd/wlroots.net) — incomplete .NET wlroots bindings
