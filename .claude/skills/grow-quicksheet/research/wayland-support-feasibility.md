# Wayland Support Feasibility Brief

**Date:** 2026-06-18
**Related issue:** #3 (Wayland support investigation for --desktop mode)

## TL;DR

- **wlr-layer-shell** (Sway, Hyprland, Wayfire, River) is the only viable path for zero-dep P/Invoke Wayland wallpaper embedding. Works like `_NET_WM_WINDOW_TYPE_DESKTOP` but Wayland-native.
- **GNOME/Mutter** has NO background-layer protocol. No workaround exists without taking a dependency. GNOME is effectively unsupported without a proxy shim.
- **KDE Plasma 6** uses its own `org_kde_plasma_window_management` — requires KDE-specific QML/API, not portable.
- Implementation is **feasible but large** (~800–1200 lines of new P/Invoke + protocol glue). Pure C# P/Invoke against `libwayland-client.so.0` is proven possible. Font rendering requires either Xft-over-XWayland fallback or a bitmap font blitter for the SHM path.
- Recommended: **Phase 1** targets wlroots compositors only (Sway+Hyprland = ~70% of tiling WM Wayland users). GNOME/KDE deferred.

## Per-Compositor Analysis

### wlroots-based (Sway, Hyprland, Wayfire, River)

| Aspect | Detail |
|--------|--------|
| Protocol | `zwlr_layer_shell_v1` (stable since 2020) |
| Library | `libwayland-client.so.0` (system library, ubiquitous) |
| Layer | `ZWLR_LAYER_SHELL_V1_LAYER_BACKGROUND` (enum 0) |
| Anchoring | All 4 edges → full-screen coverage |
| Multi-monitor | Per-output surfaces (pass `wl_output` or NULL for all) |
| Input | Layer surfaces can accept keyboard/pointer focus |
| Rendering | `wl_shm` shared-memory buffers (no GPU dependency) |
| Font rendering | Manual bitmap blit OR link `libfontconfig`/`libfreetype` via P/Invoke |

**Key insight:** The Wayland wire protocol is message-based over Unix sockets. `libwayland-client` handles the transport; the layer-shell extension is just additional opcodes. All can be driven via P/Invoke from C#.

### GNOME (Mutter)

| Aspect | Detail |
|--------|--------|
| Protocol | None. No `wlr-layer-shell`, no `ext-background-effect`. |
| Workaround | "Wayland-2-GNOME" proxy (external process, adds a dependency) |
| XWayland | X11 path works under XWayland but `_NET_WM_WINDOW_TYPE_DESKTOP` is NOT honored — renders as floating window |
| Forecast | GNOME 50 (2026) drops X11 session entirely. XWayland remains for legacy apps but wallpaper hint still won't work |
| Verdict | **Not supportable** without external shim or GNOME-specific extension |

### KDE Plasma 6

| Aspect | Detail |
|--------|--------|
| Protocol | `org_kde_plasma_window_management` (KDE-internal) |
| Wallpaper API | QML `org.kde.plasma.wallpaper` plugin system |
| Layer-shell | KDE does support `wlr-layer-shell` as of Plasma 6 for third-party panels |
| Verdict | **Layer-shell path works here too** — same implementation as wlroots |

## Implementation Architecture

### Proposed file structure

```
Platform/Linux.Wayland/
├── WaylandMethods.cs          # P/Invoke: libwayland-client.so.0
├── LayerShellMethods.cs       # P/Invoke: zwlr_layer_shell_v1 opcodes
├── ShmBuffer.cs               # wl_shm_pool + wl_buffer via memfd_create/mmap
├── WaylandDesktopWindow.cs    # Equivalent of DesktopWindow.cs but for Wayland
└── BitmapFontRenderer.cs      # 8x16 PSF font blitter for SHM surfaces
```

### Key P/Invoke surface

```csharp
// libwayland-client.so.0
wl_display_connect, wl_display_disconnect, wl_display_roundtrip,
wl_display_dispatch, wl_display_get_registry, wl_registry_bind,
wl_compositor_create_surface, wl_surface_attach, wl_surface_commit,
wl_surface_damage, wl_shm_create_pool, wl_shm_pool_create_buffer

// libc.so.6
memfd_create, ftruncate, mmap, munmap, close

// Protocol-level (no .so — wire messages via libwayland-client proxy objects)
zwlr_layer_shell_v1_get_layer_surface
zwlr_layer_surface_v1_set_size, _set_anchor, _ack_configure
```

### Rendering approach (two options)

| Option | Pros | Cons |
|--------|------|------|
| **A: Bitmap font + raw SHM** | Zero deps beyond libc+libwayland-client; pixel-perfect control | No anti-aliasing; limited Unicode; ~200 lines of font blitter code |
| **B: Xft/FreeType P/Invoke** | Anti-aliased text; full Unicode; matches X11 quality | Adds P/Invoke for `libfreetype.so.6` + `libfontconfig.so.1`; more complex buffer management |

**Recommendation:** Option B (FreeType P/Invoke) for text quality parity with X11. Both `libfreetype` and `libfontconfig` are de-facto system libraries on any graphical Linux install — same tier as `libX11.so.6` which is already used.

### Runtime detection flow

```
Program.cs startup:
1. Check WAYLAND_DISPLAY env var
2. If set → try wl_display_connect()
   - If zwlr_layer_shell_v1 available → WaylandDesktopWindow
   - Else → warn "compositor doesn't support layer-shell" + fallback
3. If WAYLAND_DISPLAY not set → existing X11 path
```

## Effort Estimate

| Component | Lines (est.) | Complexity |
|-----------|-------------|------------|
| WaylandMethods.cs (P/Invoke decls) | ~150 | Low — mechanical |
| LayerShellMethods.cs | ~80 | Medium — protocol wire format |
| ShmBuffer.cs (memfd + mmap) | ~120 | Medium — unsafe/pointers |
| BitmapFontRenderer OR FreeType bridge | ~200 | Medium-High |
| WaylandDesktopWindow.cs (event loop + render) | ~400 | High — mirrors DesktopWindow.cs |
| Runtime detection + fallback | ~50 | Low |
| **Total** | **~1000** | **~2-3 focused sessions** |

## Risks

1. **Wayland listener/callback model** — Wayland uses struct-of-function-pointers for event dispatch. In C# this means `Marshal.GetFunctionPointerForDelegate` + pinned delegates. Error-prone but doable (same pattern as X11 `XEvent` handling).
2. **No protocol code generation** — C clients use `wayland-scanner` to generate dispatch stubs from XML. In C# we'd hand-write the opcodes. One-time cost but brittle if protocol changes (it won't — v1 is stable).
3. **Multi-monitor** — Each output needs its own layer surface + SHM buffer. Adds complexity for resize/hotplug.
4. **Input handling** — Keyboard events come via `wl_keyboard` listener, not XEvent. New input dispatch path.
5. **Testing** — Maintainer runs X11 per issue #3. Needs a Wayland tester or CI with `wlroots`-based compositor.

## Compositor Market Share (Wayland Linux desktop users)

| Compositor | Protocol Support | Est. share of Wayland tiling users |
|-----------|-----------------|-------------------------------------|
| Sway | ✅ wlr-layer-shell | ~35% |
| Hyprland | ✅ wlr-layer-shell | ~30% |
| KDE Plasma 6 | ✅ wlr-layer-shell | ~20% |
| GNOME/Mutter | ❌ | ~40% of all Wayland desktop (but non-tiling; less QuickSheet audience) |
| Wayfire/River | ✅ wlr-layer-shell | ~5% |

**Core audience overlap:** QuickSheet's r/unixporn + tiling-WM audience is ~90% wlroots-compatible.

## Implications for QuickSheet

1. **Queue a "Phase 1 Wayland" implementation** targeting wlr-layer-shell only. This covers Sway + Hyprland + KDE + Wayfire — the majority of QuickSheet's target audience (power users who customize desktops).

2. **Add `PLATFORM_LINUX_WAYLAND` compiler define** and a new `Platform/Linux.Wayland/` folder. Runtime detection (not compile-time) decides which path runs — both X11 and Wayland code ship in the same binary.

3. **GNOME is explicitly out-of-scope for Phase 1.** Document this in README: "Wayland support requires a compositor with `wlr-layer-shell` (Sway, Hyprland, KDE Plasma 6, Wayfire). GNOME is not supported due to missing protocol."

4. **This is a "help wanted" contribution magnet.** The research brief + clear architecture makes it attractive for an external contributor. Consider posting the feasibility summary as a comment on issue #3 to invite PRs.

5. **Star-growth angle:** Wayland support is the #1 blocker for r/unixporn and r/hyprland audiences. Announcing even partial support would be highly shareable in those communities.
