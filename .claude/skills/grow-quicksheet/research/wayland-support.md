# Wayland Support Investigation for QuickSheet Desktop Mode

> Research brief for [Issue #3](https://github.com/cemheren/QuickSheet/issues/3)

## TL;DR

- **wlr-layer-shell** (background layer) is the correct protocol for Sway/Hyprland/wlroots compositors. Covers ~80% of tiling-WM Wayland users.
- GNOME and KDE have **no public wallpaper-embedding protocol** for third-party apps. GNOME is effectively unreachable; KDE's portal is unstable/not merged.
- Hand-written P/Invoke against `libwayland-client.so.0` is feasible but extremely tedious — Wayland's object model requires vtable structs and callback marshalling for every interface.
- A **thin C shim library** (~300 lines) wrapping `wl_display_connect` → `zwlr_layer_shell_v1_get_layer_surface` → SHM buffer attach is the pragmatic path. QuickSheet P/Invokes the shim, not raw libwayland.
- Runtime detection: check `$WAYLAND_DISPLAY` vs `$DISPLAY` to choose code path.

---

## Compositor Support Matrix

| Compositor | Protocol | Status | Notes |
|---|---|---|---|
| **Sway** | `wlr-layer-shell-unstable-v1` | ✅ Mature | Background layer works perfectly |
| **Hyprland** | `wlr-layer-shell-unstable-v1` | ✅ Full | Same protocol, fully supported |
| **River** | `wlr-layer-shell-unstable-v1` | ✅ | wlroots-based |
| **Wayfire** | `wlr-layer-shell-unstable-v1` | ✅ | wlroots-based |
| **GNOME** | None (proprietary shell) | ❌ | No third-party background layer. XWayland fallback renders as regular window. |
| **KDE Plasma 6** | `org_kde_plasma_window_management` (private) | ❌ | Not exposed to third-party. Portal API for wallpaper-*image* only, not embedded surface. |

**Conclusion:** Target wlroots compositors only (Sway + Hyprland + River + Wayfire). These are exactly the r/unixporn audience — highest-value for stars anyway.

---

## Implementation Approaches

### Option A: Pure P/Invoke against `libwayland-client.so.0` (No NuGet)

**Pros:** Zero external deps (matches project policy).
**Cons:** Enormous boilerplate. Wayland's C API uses `wl_interface` structs with method-count/event-count/method-signature arrays. Each protocol object (display, registry, compositor, shm, layer-shell, surface, buffer) needs its own interface struct + listener delegate struct + method wrappers.

Estimated effort: ~800-1200 lines of interop code for a minimal working surface.

Key P/Invoke signatures needed:
```csharp
// libwayland-client.so.0
wl_display_connect(string?) → IntPtr
wl_display_disconnect(IntPtr)
wl_display_dispatch(IntPtr) → int
wl_display_roundtrip(IntPtr) → int
wl_display_get_registry(IntPtr) → IntPtr
wl_proxy_marshal_flags(...) → IntPtr  // generic message dispatch
wl_proxy_add_listener(IntPtr, IntPtr, IntPtr) → int

// The tricky part: wl_interface structs must be constructed in managed memory
// with exact binary layout matching libwayland's expectations.
```

The layer-shell protocol adds: `zwlr_layer_shell_v1_get_layer_surface`, `zwlr_layer_surface_v1_set_size`, `zwlr_layer_surface_v1_set_anchor`, `zwlr_layer_surface_v1_set_exclusive_zone`, commit/ack_configure callbacks.

### Option B: Thin C Shim Library (`libqs-wayland.so`) — RECOMMENDED

Write a ~300-line C file that:
1. Connects to Wayland display
2. Binds `wl_compositor`, `wl_shm`, `zwlr_layer_shell_v1`
3. Creates a layer-shell surface on the background layer (anchored all edges, exclusive zone -1)
4. Creates an SHM buffer pool
5. Exposes a simple API to C#:

```c
// libqs-wayland.h — public API for QuickSheet
typedef struct QsWayland QsWayland;

QsWayland* qs_wayland_init(int width, int height);
void*      qs_wayland_get_buffer(QsWayland* ctx);   // returns ARGB pixel pointer
void       qs_wayland_commit(QsWayland* ctx);       // wl_surface_commit
int        qs_wayland_dispatch(QsWayland* ctx);     // single event loop iteration
void       qs_wayland_destroy(QsWayland* ctx);

// Input events delivered via callback registration
typedef void (*QsKeyCallback)(uint32_t key, uint32_t state, void* data);
void qs_wayland_set_key_callback(QsWayland* ctx, QsKeyCallback cb, void* data);
```

C# side: P/Invoke these 6 functions. Render into the pixel buffer exactly like the current X11 path renders via Xlib/Xft — except drawing to a memory buffer instead of an X11 drawable.

**Build:** `cc -shared -o libqs-wayland.so qs-wayland.c $(pkg-config --cflags --libs wayland-client)` — compiled on the user's machine or shipped as a prebuilt `.so` in releases.

**Pros:**
- Zero NuGet deps (C shim is project source, not a package)
- Clean separation: all Wayland complexity in one C file
- QuickSheet C# code only needs ~50 lines of new P/Invoke + a rendering adapter
- Matches the existing X11 pattern (P/Invoke native drawing libs)

**Cons:**
- Requires `wayland-client` dev headers to compile (standard on all Wayland systems)
- Adds a build step for the shim (or distribute prebuilt)
- Requires `wayland-scanner` to generate protocol headers from XML

### Option C: Existing NuGet Bindings (WaylandSharp / WaylandDotnet)

**Ruled out** by zero-NuGet-deps policy. Noted for completeness:
- WaylandSharp (0.2.1): .NET 6+, source generator from protocol XML. Stale (last update 2022).
- WaylandDotnet: .NET 10+, LibraryImport-based. Active but too new.

---

## Architecture Proposal

```
Platform/Linux/
├── LinuxDesktopHost.cs         (existing — runtime detect X11 vs Wayland)
├── DesktopWindow.cs            (existing X11 path — unchanged)
├── X11Methods.cs               (existing)
├── WaylandDesktopHost.cs       (NEW — P/Invokes libqs-wayland.so)
└── Wayland/
    ├── qs-wayland.c            (NEW — C shim, ~300 lines)
    ├── qs-wayland.h            (NEW)
    └── Makefile                (NEW — optional, for local build)
```

**Runtime detection in `LinuxDesktopHost.cs`:**
```csharp
public void Run(string? csvPath)
{
    if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("WAYLAND_DISPLAY"))
        && File.Exists("/usr/lib/libqs-wayland.so"))  // or bundled path
    {
        var waylandHost = new WaylandDesktopHost();
        waylandHost.Run(csvPath);
    }
    else
    {
        // Existing X11 path
        _display = XOpenDisplay(null);
        ...
    }
}
```

**Rendering strategy:** The current X11 path uses Xlib drawing primitives (XDrawRectangle, XDrawString via Xft). The Wayland path would:
1. Get a raw ARGB buffer from the shim
2. Render text using a software rasterizer (options: hand-roll bitmap font, or P/Invoke `libfreetype.so` + `libfontconfig.so` for TTF — both are standard system libs)
3. Commit buffer

**Simpler alternative for text:** Use `libcairo.so` (universally installed on Linux desktops) as the rendering backend for the Wayland path. P/Invoke `cairo_image_surface_create_for_data` pointing at the SHM buffer, then use Cairo's text API. Cairo is a system library, not a NuGet package — same category as libX11/libXft.

---

## Input Handling

The layer-shell surface can receive keyboard focus (set `keyboard_interactivity = ON_DEMAND`). The C shim handles `wl_keyboard` events and forwards via callback. Mouse clicks work similarly via `wl_pointer`.

On Sway: background-layer surfaces don't get keyboard focus by default. Users would need `swaymsg` focus commands or the surface uses `EXCLUSIVE` keyboard interactivity (compositor may deny). This matches X11 behavior where desktop-type windows also need special focus handling.

**Practical note:** Most QuickSheet users on tiling WMs (the primary Wayland audience) will interact via keybinds that focus the window, same as they do now with X11.

---

## Effort Estimate

| Component | Lines | Difficulty |
|---|---|---|
| C shim (`qs-wayland.c`) | ~250-350 | Medium (standard Wayland client boilerplate) |
| C# P/Invoke wrapper | ~50-80 | Low |
| WaylandDesktopHost.cs (render loop) | ~200-300 | Medium (port grid rendering to buffer) |
| Cairo text rendering adapter | ~100-150 | Low (if using Cairo) |
| Runtime detection logic | ~20 | Trivial |
| **Total** | **~650-900** | Medium overall |

---

## Phased Rollout

1. **Phase 1:** C shim + basic colored-rectangle rendering (no text). Proves layer-shell works.
2. **Phase 2:** Cairo text rendering → full grid with cell content.
3. **Phase 3:** Input handling (keyboard + mouse click on cells).
4. **Phase 4:** Feature parity with X11 path (editing, search, extensions).

Phase 1 alone is a valid "Wayland support (experimental)" PR that unblocks testing.

---

## Implications for QuickSheet

1. **Queue Bucket E action:** Implement Phase 1 (C shim + WaylandDesktopHost skeleton). Small enough for one run if shim is kept minimal. Opens a PR that partially addresses #3.
2. **Queue Bucket A action:** Add "Wayland (Sway/Hyprland)" to README's platform support section once Phase 1 merges. This is a discoverability win — "Wayland support" is a top search term for Linux desktop tools.
3. **Queue Bucket C action:** Draft an r/swaywm or r/hyprland post showcasing QuickSheet as a Wayland-native desktop wallpaper spreadsheet (once Phase 2+ works). This audience is extremely receptive to novel desktop customization.

---

## References

- [wlr-layer-shell-unstable-v1 protocol XML](https://github.com/swaywm/wlr-protocols/blob/master/unstable/wlr-layer-shell-unstable-v1.xml)
- [wf-background](https://github.com/WayfireWM/wf-background) — minimal C layer-shell wallpaper implementation
- [Sway layer-shell support](https://www.phoronix.com/news/Sway-1.9-rc1-Wayland-Compositor)
- [Hyprland protocol support](https://hypr.land/)
- [WaylandSharp](https://github.com/X9VoiD/WaylandSharp) (NuGet — ruled out by policy)
- [Wayland SHM buffer example](https://github.com/wayland-project/wayland/blob/main/clients/simple-shm.c)
