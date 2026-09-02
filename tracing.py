"""Making a UIA run watchable, without touching the user's mouse.

Driving controls through UIA patterns is invisible -- fields fill themselves
with no pointer motion, so a human watching cannot tell a working run from a
stalled one, and a wrong-control bug looks exactly like a right one.

So we draw our own pointer instead of moving the real one: a click-through,
always-on-top overlay, painted from its own thread, showing a gradient arrow
gliding to each control and a glowing box on the control it is about to act
on. Set SHOW_TRAILS to keep the arcs it travelled on screen as well, so
repeated hops between the same two controls bundle into loops and the run's
whole path through the form is readable at a glance. The real mouse and
keyboard stay yours throughout.

The overlay is a per-pixel-alpha layered window fed by Pillow rather than a Tk
canvas: the look wanted here is soft glow and translucency, which a colour-key
window cannot do (it can only cut pixels fully in or fully out), and Tk's
polygons are aliased besides. We repaint only the bounding box of what is on
screen and move the window to it, so a frame costs a small image rather than a
full-screen one.

Off by default so unattended runs stay fast; call `configure(visual=True)` (or
set FAKTURAMA_TRACE=visual) to watch.
"""

import ctypes
import math
import os
import queue
import threading
import time
from contextlib import contextmanager
from ctypes import wintypes
from functools import lru_cache
from typing import Any, Iterator


#: Seconds the highlight stays on a control before the action fires. Long
#: enough to follow by eye, short enough not to dominate the run.
DWELL = 0.35

#: How long the cursor takes to travel to its target, and the frame budget.
GLIDE_SECONDS = 0.5
FRAME_MS = 16

#: Whether the travelled arcs and their end dots are drawn behind the cursor.
#: Off: the pointer alone, with only the highlight box marking its target.
SHOW_TRAILS = False

#: How many past arcs stay on screen when they are shown. Enough to say where
#: the run has been; more than this and the loops read as clutter.
TRAIL_KEEP = 7

#: How far an arc bows out from the straight line, as a fraction of its
#: length, and how much that grows per trip. Successive hops between the same
#: pair of controls fan out into the loop bundle instead of overdrawing.
BOW = 0.12
BOW_STEP = 0.05

#: Palette, RGBA. Cool trails and a dark warm cursor, so the pointer stays
#: legible against both the trails and whatever is underneath.
TRAIL_GLOW = (198, 228, 255)
TRAIL_CORE = (255, 255, 255)
NODE = (255, 255, 255)
NODE_ACCENT = (255, 138, 208)
HIGHLIGHT_RGB = (255, 110, 196)
#: Second highlight colour, for "this is the thing we were waiting for".
CONFIRM = (150, 240, 200)
#: The pointer's iridescent fill, sampled along its axis: cyan into white
#: into pink, with a bright rim to lift it off dark backgrounds.
CURSOR_A = (34, 216, 255)
CURSOR_B = (255, 255, 255)
CURSOR_C = (255, 133, 203)
CURSOR_RIM = (255, 255, 255, 165)

#: Cursor height in pixels, and the blur radius used for every glow pass.
CURSOR_SIZE = 26

#: Where the arrow points when it is not travelling, in `_rotate` degrees.
CURSOR_REST_ANGLE = -38.0
#: Turning is deliberately lazy: the arrow leans into the direction of travel
#: over several frames rather than snapping, which is what makes it read as a
#: thing being steered. Fraction of the remaining turn per frame, and a hard
#: cap in degrees per frame.
CURSOR_TURN = 0.20
CURSOR_TURN_MAX = 8.0
#: Rotations are rendered to this granularity, so the glyph cache stays small.
CURSOR_ANGLE_STEP = 5.0
GLOW_RADIUS = 6

_visual = os.environ.get("FAKTURAMA_TRACE") == "visual"
_log = True
_depth = 0
_overlay: "_Overlay | None" = None


def configure(visual: bool | None = None, log: bool | None = None) -> None:
    """Turn the on-screen pointer and the step log on or off.

    Args:
        visual (bool | None): Draw the overlay cursor, arcs and highlight.
        log (bool | None): Print one line per step.
    """
    global _visual, _log
    if visual is not None:
        _visual = visual
    if log is not None:
        _log = log


@contextmanager
def step(label: str) -> Iterator[None]:
    """Log one automation step and how long it took.

    Args:
        label (str): What the step does, in UI terms ("set Date to 2026-04-17").
    """
    global _depth
    if _log:
        print(f"{'  ' * _depth}-> {label}", flush=True)
    _depth += 1
    start = time.perf_counter()
    try:
        yield
    except BaseException as exc:
        if _log:
            print(f"{'  ' * (_depth - 1)}!! {label}: {type(exc).__name__}: {exc}", flush=True)
        raise
    finally:
        _depth -= 1
    if _log:
        print(f"{'  ' * _depth}   {label} ({time.perf_counter() - start:.2f}s)", flush=True)


def point_at(ctrl: Any, colour: tuple[int, int, int] = HIGHLIGHT_RGB) -> None:
    """Glide the drawn cursor onto `ctrl` and box it.

    Purely cosmetic: the caller performs the real action through a UIA
    pattern, so this cannot misdirect it -- and neither can the user's own
    mouse, which is untouched.

    Args:
        ctrl (Any): The control about to be acted on.
        colour (tuple[int, int, int]): Highlight RGB.
    """
    if not _visual:
        return
    rect = ctrl.rectangle()
    overlay = _get_overlay()
    mid = rect.mid_point()
    overlay.glide_to(mid.x, mid.y)
    overlay.highlight((rect.left, rect.top, rect.right, rect.bottom), colour)
    time.sleep(DWELL)
    overlay.highlight(None)


def stop() -> None:
    """Tear the overlay down. Safe to call when it was never started."""
    global _overlay
    if _overlay is not None:
        _overlay.stop()
        _overlay = None


def _enable_dpi_awareness() -> None:
    """Put the process in physical pixels, the units UIA reports.

    A DPI-unaware process sees a virtual screen scaled down by the display
    scaling (1536x864 on a 1920x1080 screen at 125%), and Windows then scales
    its windows back up. Drawing a control's UIA rectangle in those units puts
    the pointer at 80% of the way to the control -- close enough to look
    almost right, which is the worst kind of wrong. Must run before any window
    is created; it is a process-wide switch, so we set the modern per-monitor
    mode and fall back for older Windows.
    """
    user32 = ctypes.windll.user32
    try:
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):  # per-monitor v2
            return
    except AttributeError:
        pass
    user32.SetProcessDPIAware()


def _get_overlay() -> "_Overlay":
    global _overlay
    if _overlay is None:
        _overlay = _Overlay()
        _overlay.start()
    return _overlay


# --- geometry ------------------------------------------------------------


def _arc(start: tuple[float, float], end: tuple[float, float], bow: float, steps: int = 48) -> list[tuple[float, float]]:
    """Points along a quadratic curve from `start` to `end`, bowed sideways.

    Args:
        start (tuple[float, float]): Where the cursor is.
        end (tuple[float, float]): Where it is going.
        bow (float): Sideways offset of the control point, as a fraction of
            the straight-line distance. Signed, so alternating trips bow to
            opposite sides.
        steps (int): Points to sample.

    Returns:
        list[tuple[float, float]]: The curve, start and end included.
    """
    (x0, y0), (x1, y1) = start, end
    dx, dy = x1 - x0, y1 - y0
    length = math.hypot(dx, dy) or 1.0
    # Control point: the midpoint pushed along the perpendicular.
    cx = (x0 + x1) / 2 - dy / length * length * bow
    cy = (y0 + y1) / 2 + dx / length * length * bow
    return [
        (
            (1 - t) ** 2 * x0 + 2 * (1 - t) * t * cx + t**2 * x1,
            (1 - t) ** 2 * y0 + 2 * (1 - t) * t * cy + t**2 * y1,
        )
        for t in (i / steps for i in range(steps + 1))
    ]


def _rotate(points: list[tuple[float, float]], degrees: float) -> list[tuple[float, float]]:
    """Rotate `points` about the origin, anticlockwise on screen."""
    angle = math.radians(degrees)
    cos, sin = math.cos(angle), math.sin(angle)
    return [(x * cos + y * sin, -x * sin + y * cos) for x, y in points]


def _heading(delta: tuple[float, float]) -> float:
    """The glyph rotation that makes the arrow point along `delta`.

    The arrow is modelled pointing straight up, and `_rotate` takes (0, -1) to
    (-sin d, -cos d), so we solve that for d.

    Args:
        delta (tuple[float, float]): A direction of travel, in screen pixels.

    Returns:
        float: Degrees, matching `_rotate`'s convention.
    """
    dx, dy = delta
    return math.degrees(math.atan2(-dx, -dy))


@lru_cache(maxsize=256)
def _cursor_glyph(size: int, angle: float = CURSOR_REST_ANGLE) -> tuple[Any, tuple[int, int]]:
    """Render the pointer: a rounded arrow filled with an iridescent gradient.

    Built once per (size, angle) and pasted each frame -- rebuilding this
    (supersampled polygon, gradient, blurs) at 60fps would cost more than the
    rest of the overlay put together. Callers quantise `angle` to
    CURSOR_ANGLE_STEP so a run only ever needs a few dozen of these.

    Args:
        size (int): Height of the glyph in pixels.
        angle (float): Rotation in degrees; the arrow points up at 0.

    Returns:
        tuple[Any, tuple[int, int]]: The RGBA image, and the pixel offset of
            the arrow's tip within it -- what we line up with the target.
    """
    from PIL import Image, ImageDraw, ImageFilter

    # The arrow as a concave quad: tip, wing, notch, wing, pointing up.
    shape = _rotate([(0.0, -1.0), (0.62, 0.78), (0.0, 0.30), (-0.62, 0.78)], angle)
    pad = size * 0.5
    span = size / 2
    centre = (pad + span, pad + span)
    box = int(pad * 2 + span * 2)
    # Anchor on the arrow's tip, so the tip lands on the target and the glyph
    # pivots about it while turning.
    tip = (centre[0] + shape[0][0] * span, centre[1] + shape[0][1] * span)

    # Draw the outline fat and round-jointed at 4x, then downsample: it both
    # rounds the corners and gives us an antialiased mask for free.
    ss = 4
    radius = max(2, int(size * 0.13)) * ss
    mask = Image.new("L", (box * ss, box * ss), 0)
    md = ImageDraw.Draw(mask)
    pts = [((centre[0] + x * span) * ss, (centre[1] + y * span) * ss) for x, y in shape]
    md.polygon(pts, fill=255)
    md.line(pts + [pts[0]], fill=255, width=radius * 2, joint="curve")
    for px, py in pts:
        md.ellipse((px - radius, py - radius, px + radius, py + radius), fill=255)
    mask = mask.resize((box, box), Image.LANCZOS)

    # Iridescent fill: cyan into white into pink, along the pointer's axis.
    gradient = Image.new("RGBA", (box, box))
    stops = [(0.0, CURSOR_A), (0.55, CURSOR_B), (1.0, CURSOR_C)]
    pixels = gradient.load()
    # Sample across the glyph itself, not the padded canvas: normalising over
    # the whole box leaves the arrow sitting in the middle of the ramp, which
    # renders it almost entirely white.
    lo, hi = pad, box - pad
    for y in range(box):
        for x in range(box):
            t = min(1.0, max(0.0, ((x - lo) + (y - lo)) / (2 * (hi - lo))))
            for (t0, c0), (t1, c1) in zip(stops, stops[1:]):
                if t <= t1 or t1 == 1.0:
                    k = (t - t0) / (t1 - t0) if t1 > t0 else 0.0
                    k = min(1.0, max(0.0, k))
                    pixels[x, y] = tuple(round(a + (b - a) * k) for a, b in zip(c0, c1)) + (255,)
                    break

    glyph = Image.new("RGBA", (box, box), (0, 0, 0, 0))
    # Drop shadow first, so the pointer stays readable on a light background.
    shadow = mask.filter(ImageFilter.GaussianBlur(size * 0.09)).point(lambda v: int(v * 0.45))
    glyph.paste(Image.new("RGBA", (box, box), (40, 30, 60, 255)), (0, int(size * 0.06)), shadow)
    glyph.paste(gradient, (0, 0), mask)
    # Thin bright rim: the mask minus an eroded copy of itself.
    from PIL import ImageChops

    inner = mask.filter(ImageFilter.MinFilter(3)).filter(ImageFilter.MinFilter(3))
    ring = ImageChops.subtract(mask, inner).point(lambda v: int(v * 0.85))
    glyph.paste(Image.new("RGBA", (box, box), CURSOR_RIM), (0, 0), ring)
    return glyph, (round(tip[0]), round(tip[1]))


# --- overlay -------------------------------------------------------------


class _Overlay:
    """A click-through layered window drawing the cursor, arcs and highlight.

    The window and all its GDI objects belong to one thread; the automation
    thread only posts commands onto a queue. Everything from `_run` down runs
    on the overlay thread.
    """

    def __init__(self) -> None:
        self._commands: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()

    def start(self) -> None:
        """Spawn the overlay thread and wait for the window to exist."""
        self._thread = threading.Thread(target=self._run, name="tracing-overlay", daemon=True)
        self._thread.start()
        self._ready.wait(timeout=10)

    def stop(self) -> None:
        self._commands.put(("stop", None))
        if self._thread is not None:
            self._thread.join(timeout=3)

    def glide_to(self, x: int, y: int) -> None:
        """Move the drawn cursor to (x, y), returning once it has arrived."""
        arrived = threading.Event()
        self._commands.put(("glide", (x, y, arrived)))
        arrived.wait(timeout=GLIDE_SECONDS + 3)

    def highlight(
        self, rect: tuple[int, int, int, int] | None, colour: tuple[int, int, int] = HIGHLIGHT_RGB
    ) -> None:
        """Box `rect` (screen coordinates), or clear the box when None."""
        self._commands.put(("highlight", (rect, colour)))

    # --- overlay thread ---------------------------------------------------

    def _run(self) -> None:
        _enable_dpi_awareness()
        self._canvas = _LayeredWindow()
        self._pos: tuple[float, float] = (0.0, 0.0)
        self._placed = False       # no cursor drawn until the first glide
        self._trails: list[list[tuple[float, float]]] = []
        self._current: list[tuple[float, float]] | None = None
        self._nodes: list[tuple[float, float]] = []
        self._box: tuple[int, int, int, int] | None = None
        self._box_colour = HIGHLIGHT_RGB
        self._bow_side = 1.0
        self._bow = BOW
        self._angle = CURSOR_REST_ANGLE
        self._ready.set()

        while True:
            frame_start = time.perf_counter()
            if not self._pump():
                self._canvas.destroy()
                return
            self._canvas.pump_messages()
            self._paint()
            time.sleep(max(0.0, FRAME_MS / 1000 - (time.perf_counter() - frame_start)))

    def _pump(self) -> bool:
        """Handle queued commands and advance any animation. False to stop."""
        try:
            while True:
                name, payload = self._commands.get_nowait()
                if name == "stop":
                    return False
                if name == "glide":
                    self._begin_glide(*payload)
                elif name == "highlight":
                    self._box, self._box_colour = payload
        except queue.Empty:
            pass
        self._advance()
        return True

    def _begin_glide(self, x: int, y: int, arrived: threading.Event) -> None:
        if not self._placed:
            # First move of the run: start from the target so we do not draw a
            # meaningless arc in from the top-left corner of the screen.
            self._pos = (float(x), float(y))
            self._placed = True
            arrived.set()
            return
        self._path = _arc(self._pos, (float(x), float(y)), self._bow * self._bow_side)
        self._frame = 0
        self._frames = max(1, int(GLIDE_SECONDS * 1000 / FRAME_MS))
        self._arrived = arrived
        self._current = []
        # Alternate sides and widen a little each trip, so repeated hops
        # between two controls fan into a bundle instead of one thick line.
        self._bow_side *= -1
        self._bow = BOW if self._bow > BOW + BOW_STEP * 3 else self._bow + BOW_STEP
        self._nodes.append(self._pos)
        self._nodes.append((float(x), float(y)))
        self._nodes = self._nodes[-2 * TRAIL_KEEP :]

    def _advance(self) -> None:
        """Step the in-flight glide by one frame."""
        if self._current is None:
            return
        self._frame += 1
        # Ease-out: quick off the mark, settling onto the target, which reads
        # as deliberate rather than mechanical.
        t = 1 - (1 - self._frame / self._frames) ** 3
        index = min(len(self._path) - 1, int(t * (len(self._path) - 1)))
        self._current = self._path[: index + 1]
        self._pos = self._path[index]
        # Steer towards the tangent of the arc, so the arrow leans through the
        # curve and is already facing the target when it lands.
        ahead = self._path[min(index + 2, len(self._path) - 1)]
        self._turn_towards((ahead[0] - self._pos[0], ahead[1] - self._pos[1]))
        if self._frame >= self._frames:
            self._trails.append(self._path)
            self._trails = self._trails[-TRAIL_KEEP:]
            self._current = None
            self._arrived.set()

    def _turn_towards(self, delta: tuple[float, float]) -> None:
        """Rotate the arrow a little further towards heading `delta`.

        Eased and capped rather than set outright: a pointer that snaps to
        each new bearing looks like it teleported, one that swings round over
        a few frames looks like it is being flown.

        Args:
            delta (tuple[float, float]): Direction of travel, in pixels.
        """
        if abs(delta[0]) < 0.01 and abs(delta[1]) < 0.01:
            return
        # Shortest way round: an unwrapped difference sends the arrow the long
        # way whenever a turn crosses the +/-180 boundary.
        remaining = (_heading(delta) - self._angle + 180) % 360 - 180
        move = max(-CURSOR_TURN_MAX, min(CURSOR_TURN_MAX, remaining * CURSOR_TURN))
        self._angle += move

    def _paint(self) -> None:
        from PIL import Image, ImageDraw, ImageFilter

        shapes = list(self._trails) + ([self._current] if self._current else []) if SHOW_TRAILS else []
        nodes = self._nodes if SHOW_TRAILS else []
        points = [p for shape in shapes for p in shape] + nodes
        if self._placed:
            points.append(self._pos)
        if self._box:
            left, top, right, bottom = self._box
            points += [(left, top), (right, bottom)]
        if not points:
            self._canvas.hide()
            return

        # The glyph image is twice the cursor height and anchored on its tip,
        # so the box has to allow a whole glyph beyond the cursor in any
        # direction -- `alpha_composite` refuses a paste that would overhang.
        pad = CURSOR_SIZE * 2 + GLOW_RADIUS * 3
        x0 = int(min(p[0] for p in points) - pad)
        y0 = int(min(p[1] for p in points) - pad)
        x1 = int(max(p[0] for p in points) + pad)
        y1 = int(max(p[1] for p in points) + pad)
        size = (max(1, x1 - x0), max(1, y1 - y0))

        def local(p: tuple[float, float]) -> tuple[float, float]:
            return (p[0] - x0, p[1] - y0)

        glow = Image.new("RGBA", size, (0, 0, 0, 0))
        gd = ImageDraw.Draw(glow)
        for i, shape in enumerate(shapes):
            # Older arcs fade, so the newest trip reads as the current one.
            fade = (i + 1) / len(shapes)
            gd.line([local(p) for p in shape], fill=TRAIL_GLOW + (int(70 * fade),), width=11, joint="curve")
        for node in nodes:
            nx, ny = local(node)
            gd.ellipse((nx - 9, ny - 9, nx + 9, ny + 9), fill=NODE + (60,))
        if self._box:
            left, top = local((self._box[0], self._box[1]))
            right, bottom = local((self._box[2], self._box[3]))
            gd.rounded_rectangle(
                (left - 3, top - 3, right + 3, bottom + 3),
                radius=8, outline=self._box_colour + (140,), width=7,
            )
        image = glow.filter(ImageFilter.GaussianBlur(GLOW_RADIUS))

        draw = ImageDraw.Draw(image)
        for i, shape in enumerate(shapes):
            fade = (i + 1) / len(shapes)
            draw.line(
                [local(p) for p in shape],
                fill=TRAIL_CORE + (int(40 + 150 * fade),), width=2, joint="curve",
            )
        for j, node in enumerate(nodes):
            nx, ny = local(node)
            colour = NODE_ACCENT if j % 2 else NODE
            radius = 3.5 if j % 2 else 5
            draw.ellipse((nx - radius, ny - radius, nx + radius, ny + radius), fill=colour + (230,))
        if self._box:
            left, top = local((self._box[0], self._box[1]))
            right, bottom = local((self._box[2], self._box[3]))
            draw.rounded_rectangle(
                (left - 3, top - 3, right + 3, bottom + 3),
                radius=8, outline=self._box_colour + (255,), width=2,
            )
        if self._placed:
            cx, cy = local(self._pos)
            angle = round(self._angle / CURSOR_ANGLE_STEP) * CURSOR_ANGLE_STEP
            glyph, (hx, hy) = _cursor_glyph(CURSOR_SIZE, angle % 360)
            image.alpha_composite(glyph, (round(cx) - hx, round(cy) - hy))

        self._canvas.update(image, (x0, y0))


class _LayeredWindow:
    """A borderless, click-through, per-pixel-alpha window fed RGBA images.

    Created and used from one thread only. `update` both moves/resizes the
    window and blits a frame, which is how we can afford to repaint: the
    window is only ever as big as the drawing currently needs.
    """

    _CLASS_NAME = "FakturamaTracingOverlay"

    _WS_POPUP = 0x80000000
    _WS_EX_LAYERED = 0x00080000
    _WS_EX_TRANSPARENT = 0x00000020
    _WS_EX_NOACTIVATE = 0x08000000
    _WS_EX_TOOLWINDOW = 0x00000080
    _WS_EX_TOPMOST = 0x00000008
    _SW_HIDE, _SW_SHOWNA = 0, 8
    _ULW_ALPHA = 0x00000002
    _AC_SRC_OVER, _AC_SRC_ALPHA = 0x00, 0x01

    def __init__(self) -> None:
        self._user32 = ctypes.windll.user32
        self._gdi32 = ctypes.windll.gdi32
        self._visible = False
        _declare_signatures()

        # Point the class straight at DefWindowProcW by address rather than
        # wrapping it in a Python callback: Windows dispatches WM_NCCREATE
        # while we are still inside CreateWindowExW, and a Python trampoline
        # would re-enter ctypes there with an untyped pointer lParam, which
        # fails ("argument 4: OverflowError: int too long to convert") and is
        # swallowed as an ignored callback exception. Kept on the instance so
        # the thunk outlives the window.
        wndproc_type = ctypes.WINFUNCTYPE(
            ctypes.c_long, wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM
        )
        self._wndproc = wndproc_type(
            ctypes.cast(self._user32.DefWindowProcW, ctypes.c_void_p).value
        )

        cls = _WNDCLASS()
        cls.lpfnWndProc = self._wndproc
        cls.hInstance = ctypes.windll.kernel32.GetModuleHandleW(None)
        cls.lpszClassName = self._CLASS_NAME
        self._user32.RegisterClassW(ctypes.byref(cls))  # harmless if already registered

        self._hwnd = self._user32.CreateWindowExW(
            self._WS_EX_LAYERED | self._WS_EX_TRANSPARENT | self._WS_EX_NOACTIVATE
            | self._WS_EX_TOOLWINDOW | self._WS_EX_TOPMOST,
            self._CLASS_NAME, None, self._WS_POPUP,
            0, 0, 1, 1, None, None, cls.hInstance, None,
        )
        if not self._hwnd:
            raise ctypes.WinError(ctypes.get_last_error())

    def pump_messages(self) -> None:
        """Service the window's message queue so Windows keeps it alive."""
        msg = wintypes.MSG()
        while self._user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):  # PM_REMOVE
            self._user32.TranslateMessage(ctypes.byref(msg))
            self._user32.DispatchMessageW(ctypes.byref(msg))

    def update(self, image: Any, origin: tuple[int, int]) -> None:
        """Show `image` (RGBA) with its top-left corner at `origin`."""
        import numpy as np

        width, height = image.size
        # Windows wants BGRA with the colour channels already multiplied by
        # alpha; feeding it straight RGBA gives dark haloes around every glow.
        rgba = np.frombuffer(image.tobytes(), dtype=np.uint8).reshape(height, width, 4)
        alpha = rgba[:, :, 3:4].astype(np.uint16)
        bgra = np.empty_like(rgba)
        bgra[:, :, 0] = (rgba[:, :, 2] * alpha[:, :, 0] // 255).astype(np.uint8)
        bgra[:, :, 1] = (rgba[:, :, 1] * alpha[:, :, 0] // 255).astype(np.uint8)
        bgra[:, :, 2] = (rgba[:, :, 0] * alpha[:, :, 0] // 255).astype(np.uint8)
        bgra[:, :, 3] = rgba[:, :, 3]

        screen_dc = self._user32.GetDC(None)
        mem_dc = self._gdi32.CreateCompatibleDC(screen_dc)
        info = _BITMAPINFOHEADER()
        info.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
        info.biWidth = width
        info.biHeight = -height  # top-down
        info.biPlanes = 1
        info.biBitCount = 32
        info.biCompression = 0  # BI_RGB
        bits = ctypes.c_void_p()
        bitmap = self._gdi32.CreateDIBSection(
            mem_dc, ctypes.byref(info), 0, ctypes.byref(bits), None, 0
        )
        old = self._gdi32.SelectObject(mem_dc, bitmap)
        ctypes.memmove(bits, bgra.tobytes(), width * height * 4)

        blend = _BLENDFUNCTION()
        blend.BlendOp = self._AC_SRC_OVER
        blend.SourceConstantAlpha = 255
        blend.AlphaFormat = self._AC_SRC_ALPHA
        position = wintypes.POINT(origin[0], origin[1])
        size = wintypes.SIZE(width, height)
        source = wintypes.POINT(0, 0)
        self._user32.UpdateLayeredWindow(
            self._hwnd, screen_dc, ctypes.byref(position), ctypes.byref(size),
            mem_dc, ctypes.byref(source), 0, ctypes.byref(blend), self._ULW_ALPHA,
        )

        self._gdi32.SelectObject(mem_dc, old)
        self._gdi32.DeleteObject(bitmap)
        self._gdi32.DeleteDC(mem_dc)
        self._user32.ReleaseDC(None, screen_dc)

        if not self._visible:
            self._user32.ShowWindow(self._hwnd, self._SW_SHOWNA)
            self._visible = True

    def hide(self) -> None:
        if self._visible:
            self._user32.ShowWindow(self._hwnd, self._SW_HIDE)
            self._visible = False

    def destroy(self) -> None:
        self._user32.DestroyWindow(self._hwnd)
        self._hwnd = None


def _declare_signatures() -> None:
    """Give ctypes the real prototypes for the calls we make.

    Without these, ctypes assumes every argument and return value is a C int,
    which truncates 64-bit handles and rejects module handles outright
    ("argument 11: OverflowError: int too long to convert" from
    CreateWindowExW). Idempotent, so it is safe to call per window.
    """
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    kernel32 = ctypes.windll.kernel32

    kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]

    user32.CreateWindowExW.restype = wintypes.HWND
    user32.CreateWindowExW.argtypes = [
        wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID,
    ]
    user32.DestroyWindow.argtypes = [wintypes.HWND]
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.GetDC.restype = wintypes.HDC
    user32.GetDC.argtypes = [wintypes.HWND]
    user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    user32.UpdateLayeredWindow.argtypes = [
        wintypes.HWND, wintypes.HDC, ctypes.POINTER(wintypes.POINT),
        ctypes.POINTER(wintypes.SIZE), wintypes.HDC, ctypes.POINTER(wintypes.POINT),
        wintypes.DWORD, ctypes.POINTER(_BLENDFUNCTION), wintypes.DWORD,
    ]

    gdi32.CreateCompatibleDC.restype = wintypes.HDC
    gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
    gdi32.CreateDIBSection.restype = wintypes.HBITMAP
    gdi32.CreateDIBSection.argtypes = [
        wintypes.HDC, ctypes.POINTER(_BITMAPINFOHEADER), wintypes.UINT,
        ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, wintypes.DWORD,
    ]
    gdi32.SelectObject.restype = wintypes.HGDIOBJ
    gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
    gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
    gdi32.DeleteDC.argtypes = [wintypes.HDC]


class _WNDCLASS(ctypes.Structure):
    _fields_ = [
        ("style", ctypes.c_uint),
        ("lpfnWndProc", ctypes.WINFUNCTYPE(
            ctypes.c_long, wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM)),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", wintypes.HANDLE),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class _BLENDFUNCTION(ctypes.Structure):
    _fields_ = [
        ("BlendOp", ctypes.c_byte),
        ("BlendFlags", ctypes.c_byte),
        ("SourceConstantAlpha", ctypes.c_byte),
        ("AlphaFormat", ctypes.c_byte),
    ]


if __name__ == "__main__":
    # Smoke test: hop between two points so the arcs bundle into loops, then
    # a third, without ever touching the real mouse.
    configure(visual=True)
    ov = _get_overlay()
    a, b, c = (600, 700), (1300, 300), (900, 850)
    for target in [a, b, a, b, a, b, c]:
        ov.glide_to(*target)
        ov.highlight((target[0] - 90, target[1] - 26, target[0] + 90, target[1] + 26))
        time.sleep(0.35)
        ov.highlight(None)
    time.sleep(1.0)
    stop()
