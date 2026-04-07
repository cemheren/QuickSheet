using System.Runtime.InteropServices;

namespace ExcelConsole.Platform.Linux;

/// <summary>
/// P/Invoke declarations for X11 (libX11), Xft (libXft), and Xrender (libXrender).
/// These are standard system libraries on any graphical Fedora installation.
/// </summary>
internal static class X11Methods
{
    private const string LibX11 = "libX11.so.6";
    private const string LibXft = "libXft.so.2";

    // ── Display & Window ─────────────────────────────────────────────

    [DllImport(LibX11)]
    public static extern IntPtr XOpenDisplay(string? display);

    [DllImport(LibX11)]
    public static extern int XCloseDisplay(IntPtr display);

    [DllImport(LibX11)]
    public static extern IntPtr XDefaultRootWindow(IntPtr display);

    [DllImport(LibX11)]
    public static extern int XDefaultScreen(IntPtr display);

    [DllImport(LibX11)]
    public static extern int XDisplayWidth(IntPtr display, int screen);

    [DllImport(LibX11)]
    public static extern int XDisplayHeight(IntPtr display, int screen);

    [DllImport(LibX11)]
    public static extern int XDisplayWidthMM(IntPtr display, int screen);

    [DllImport(LibX11)]
    public static extern int XDisplayHeightMM(IntPtr display, int screen);

    [DllImport(LibX11)]
    public static extern IntPtr XResourceManagerString(IntPtr display);

    [DllImport(LibX11)]
    public static extern int XDefaultDepth(IntPtr display, int screen);

    [DllImport(LibX11)]
    public static extern IntPtr XDefaultVisual(IntPtr display, int screen);

    [DllImport(LibX11)]
    public static extern IntPtr XDefaultColormap(IntPtr display, int screen);

    [DllImport(LibX11)]
    public static extern IntPtr XCreateSimpleWindow(IntPtr display, IntPtr parent,
        int x, int y, uint width, uint height, uint borderWidth,
        ulong border, ulong background);

    [DllImport(LibX11)]
    public static extern IntPtr XCreateWindow(IntPtr display, IntPtr parent,
        int x, int y, uint width, uint height, uint borderWidth,
        int depth, uint @class, IntPtr visual, ulong valueMask,
        ref XSetWindowAttributes attributes);

    [DllImport(LibX11)]
    public static extern int XMatchVisualInfo(IntPtr display, int screen,
        int depth, int @class, out XVisualInfo vinfo);

    [DllImport(LibX11)]
    public static extern IntPtr XCreateColormap(IntPtr display, IntPtr window,
        IntPtr visual, int alloc);

    [DllImport(LibX11)]
    public static extern int XMapWindow(IntPtr display, IntPtr window);

    [DllImport(LibX11)]
    public static extern int XUnmapWindow(IntPtr display, IntPtr window);

    [DllImport(LibX11)]
    public static extern int XDestroyWindow(IntPtr display, IntPtr window);

    [DllImport(LibX11)]
    public static extern int XStoreName(IntPtr display, IntPtr window, string name);

    [DllImport(LibX11)]
    public static extern int XMoveResizeWindow(IntPtr display, IntPtr window,
        int x, int y, uint width, uint height);

    // ── Events ───────────────────────────────────────────────────────

    [DllImport(LibX11)]
    public static extern int XSelectInput(IntPtr display, IntPtr window, long eventMask);

    [DllImport(LibX11)]
    public static extern int XNextEvent(IntPtr display, IntPtr eventReturn);

    [DllImport(LibX11)]
    public static extern int XPending(IntPtr display);

    [DllImport(LibX11)]
    public static extern int XEventsQueued(IntPtr display, int mode);

    // ── Graphics Context ─────────────────────────────────────────────

    [DllImport(LibX11)]
    public static extern IntPtr XCreateGC(IntPtr display, IntPtr drawable,
        ulong valueMask, IntPtr values);

    [DllImport(LibX11)]
    public static extern int XFreeGC(IntPtr display, IntPtr gc);

    [DllImport(LibX11)]
    public static extern int XSetForeground(IntPtr display, IntPtr gc, ulong color);

    [DllImport(LibX11)]
    public static extern int XSetBackground(IntPtr display, IntPtr gc, ulong color);

    [DllImport(LibX11)]
    public static extern int XFillRectangle(IntPtr display, IntPtr drawable, IntPtr gc,
        int x, int y, uint width, uint height);

    [DllImport(LibX11)]
    public static extern int XDrawRectangle(IntPtr display, IntPtr drawable, IntPtr gc,
        int x, int y, uint width, uint height);

    [DllImport(LibX11)]
    public static extern int XDrawLine(IntPtr display, IntPtr drawable, IntPtr gc,
        int x1, int y1, int x2, int y2);

    [DllImport(LibX11)]
    public static extern int XClearWindow(IntPtr display, IntPtr window);

    [DllImport(LibX11)]
    public static extern int XFlush(IntPtr display);

    [DllImport(LibX11)]
    public static extern int XSync(IntPtr display, bool discard);

    // ── Color ────────────────────────────────────────────────────────

    [DllImport(LibX11)]
    public static extern int XAllocColor(IntPtr display, IntPtr colormap, ref XColor color);

    [DllImport(LibX11)]
    public static extern int XFreeColors(IntPtr display, IntPtr colormap,
        ulong[] pixels, int npixels, ulong planes);

    // ── Atoms & Properties ───────────────────────────────────────────

    [DllImport(LibX11)]
    public static extern IntPtr XInternAtom(IntPtr display, string atomName, bool onlyIfExists);

    [DllImport(LibX11)]
    public static extern int XChangeProperty(IntPtr display, IntPtr window, IntPtr property,
        IntPtr type, int format, int mode, IntPtr data, int nelements);

    [DllImport(LibX11)]
    public static extern int XChangeProperty(IntPtr display, IntPtr window, IntPtr property,
        IntPtr type, int format, int mode, IntPtr[] data, int nelements);

    // ── Clipboard (Selections) ───────────────────────────────────────

    [DllImport(LibX11)]
    public static extern IntPtr XGetSelectionOwner(IntPtr display, IntPtr selection);

    [DllImport(LibX11)]
    public static extern int XSetSelectionOwner(IntPtr display, IntPtr selection,
        IntPtr owner, ulong time);

    [DllImport(LibX11)]
    public static extern int XConvertSelection(IntPtr display, IntPtr selection,
        IntPtr target, IntPtr property, IntPtr requestor, ulong time);

    [DllImport(LibX11)]
    public static extern int XSendEvent(IntPtr display, IntPtr window, bool propagate,
        long eventMask, IntPtr eventSend);

    [DllImport(LibX11)]
    public static extern int XGetWindowProperty(IntPtr display, IntPtr window,
        IntPtr property, long longOffset, long longLength, bool delete,
        IntPtr reqType, out IntPtr actualType, out int actualFormat,
        out ulong nitems, out ulong bytesAfter, out IntPtr prop);

    [DllImport(LibX11)]
    public static extern int XFree(IntPtr data);

    // ── Keyboard ─────────────────────────────────────────────────────

    [DllImport(LibX11)]
    public static extern ulong XLookupKeysym(ref XKeyEvent keyEvent, int index);

    [DllImport(LibX11)]
    public static extern int XLookupString(ref XKeyEvent eventStruct,
        IntPtr bufferReturn, int bytesBuffer, out ulong keysymReturn, IntPtr composeStatus);

    // ── Xft (font rendering) ────────────────────────────────────────

    [DllImport(LibXft)]
    public static extern IntPtr XftFontOpenName(IntPtr display, int screen, string name);

    [DllImport(LibXft)]
    public static extern void XftFontClose(IntPtr display, IntPtr font);

    [DllImport(LibXft)]
    public static extern IntPtr XftDrawCreate(IntPtr display, IntPtr drawable,
        IntPtr visual, IntPtr colormap);

    [DllImport(LibXft)]
    public static extern void XftDrawDestroy(IntPtr draw);

    [DllImport(LibXft)]
    public static extern void XftDrawStringUtf8(IntPtr draw, ref XftColor color,
        IntPtr font, int x, int y, byte[] str, int len);

    [DllImport(LibXft)]
    public static extern void XftTextExtentsUtf8(IntPtr display, IntPtr font,
        byte[] str, int len, out XGlyphInfo extents);

    [DllImport(LibXft)]
    public static extern bool XftColorAllocValue(IntPtr display, IntPtr visual,
        IntPtr colormap, ref XRenderColor color, out XftColor result);

    [DllImport(LibXft)]
    public static extern void XftColorFree(IntPtr display, IntPtr visual,
        IntPtr colormap, ref XftColor color);

    [DllImport(LibXft)]
    public static extern void XftDrawRect(IntPtr draw, ref XftColor color,
        int x, int y, uint width, uint height);

    // ── Event masks ──────────────────────────────────────────────────

    public const long KeyPressMask = 1L << 0;
    public const long KeyReleaseMask = 1L << 1;
    public const long ButtonPressMask = 1L << 2;
    public const long ButtonReleaseMask = 1L << 3;
    public const long ExposureMask = 1L << 15;
    public const long StructureNotifyMask = 1L << 17;
    public const long FocusChangeMask = 1L << 21;

    // ── Event types ──────────────────────────────────────────────────

    public const int KeyPress = 2;
    public const int KeyRelease = 3;
    public const int ButtonPress = 4;
    public const int ButtonRelease = 5;
    public const int Expose = 12;
    public const int ConfigureNotify = 22;
    public const int SelectionClear = 29;
    public const int SelectionRequest = 30;
    public const int SelectionNotify = 31;
    public const int ClientMessage = 33;

    // ── Property change modes ────────────────────────────────────────

    public const int PropModeReplace = 0;

    // ── Key syms ─────────────────────────────────────────────────────

    public const ulong XK_Return = 0xff0d;
    public const ulong XK_Escape = 0xff1b;
    public const ulong XK_Tab = 0xff09;
    public const ulong XK_BackSpace = 0xff08;
    public const ulong XK_Delete = 0xffff;
    public const ulong XK_Left = 0xff51;
    public const ulong XK_Up = 0xff52;
    public const ulong XK_Right = 0xff53;
    public const ulong XK_Down = 0xff54;
    public const ulong XK_Home = 0xff50;
    public const ulong XK_End = 0xff57;

    // Letter keys for Ctrl+<key> handling
    public const ulong XK_a = 0x0061;
    public const ulong XK_c = 0x0063;
    public const ulong XK_d = 0x0064;
    public const ulong XK_h = 0x0068;
    public const ulong XK_o = 0x006f;
    public const ulong XK_p = 0x0070;
    public const ulong XK_q = 0x0071;
    public const ulong XK_s = 0x0073;
    public const ulong XK_v = 0x0076;
    public const ulong XK_x = 0x0078;

    // State masks
    public const uint ShiftMask = 1 << 0;
    public const uint LockMask = 1 << 1;
    public const uint ControlMask = 1 << 2;
    public const uint Mod1Mask = 1 << 3; // Alt

    // ── Structures ───────────────────────────────────────────────────

    [StructLayout(LayoutKind.Sequential)]
    public struct XColor
    {
        public ulong pixel;
        public ushort red, green, blue;
        public byte flags;
        public byte pad;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XRenderColor
    {
        public ushort red, green, blue, alpha;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XftColor
    {
        public ulong pixel;
        public XRenderColor color;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XGlyphInfo
    {
        public ushort width, height;
        public short x, y;
        public short xOff, yOff;
    }

    // XEvent is a union of 192 bytes on 64-bit
    public const int XEventSize = 192;

    // NOTE: X11's Bool is typedef'd to int (4 bytes), NOT C# bool (1 byte).
    // All structs use int for Bool fields to match the native layout.

    [StructLayout(LayoutKind.Sequential)]
    public struct XAnyEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr window;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XKeyEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr window;
        public IntPtr root;
        public IntPtr subwindow;
        public ulong time;
        public int x, y;
        public int x_root, y_root;
        public uint state;
        public uint keycode;
        public int same_screen; // Bool
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XButtonEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr window;
        public IntPtr root;
        public IntPtr subwindow;
        public ulong time;
        public int x, y;
        public int x_root, y_root;
        public uint state;
        public uint button;
        public int same_screen; // Bool
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XExposeEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr window;
        public int x, y;
        public int width, height;
        public int count;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XConfigureEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr @event;  // event window
        public IntPtr window;  // configured window
        public int x, y;
        public int width, height;
        public int border_width;
        public IntPtr above;
        public int override_redirect; // Bool
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XSelectionRequestEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr owner;
        public IntPtr requestor;
        public IntPtr selection;
        public IntPtr target;
        public IntPtr property;
        public ulong time;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XSelectionEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr requestor;
        public IntPtr selection;
        public IntPtr target;
        public IntPtr property;
        public ulong time;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XClientMessageEvent
    {
        public int type;
        public ulong serial;
        public int send_event; // Bool
        public IntPtr display;
        public IntPtr window;
        public IntPtr message_type;
        public int format;
        public long data0, data1, data2, data3, data4;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XVisualInfo
    {
        public IntPtr visual;
        public ulong visualid;
        public int screen;
        public int depth;
        public int @class;
        public ulong red_mask;
        public ulong green_mask;
        public ulong blue_mask;
        public int colormap_size;
        public int bits_per_rgb;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct XSetWindowAttributes
    {
        public IntPtr background_pixmap;
        public ulong background_pixel;
        public IntPtr border_pixmap;
        public ulong border_pixel;
        public int bit_gravity;
        public int win_gravity;
        public int backing_store;
        public ulong backing_planes;
        public ulong backing_pixel;
        public int save_under; // Bool
        public long event_mask;
        public long do_not_propagate_mask;
        public int override_redirect; // Bool
        public IntPtr colormap;
        public IntPtr cursor;
    }

    // XCreateWindow class values
    public const uint InputOutput = 1;

    // XSetWindowAttributes value mask bits
    public const ulong CWBackPixel = 1L << 1;
    public const ulong CWBorderPixel = 1L << 3;
    public const ulong CWColormap = 1L << 13;

    // XCreateColormap alloc values
    public const int AllocNone = 0;

    // Visual class for XMatchVisualInfo
    public const int TrueColor = 4;
}
