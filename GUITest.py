# -*- coding: utf-8 -*-
import sys, json, threading, time, re, csv, queue
from pathlib import Path
from collections import deque

from PyQt5.QtCore import Qt, pyqtSignal, QObject, QSize, QTimer
from PyQt5.QtWidgets import (
    QApplication, QWidget, QMainWindow, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QComboBox, QDoubleSpinBox, QSpinBox, QGroupBox, QFormLayout,
    QMessageBox, QTabWidget, QCheckBox, QPlainTextEdit, QLineEdit, QSlider, QAction, QStyleFactory
)
from PyQt5.QtGui import QPainter, QColor, QPen, QBrush, QPalette

# -------- OPTIONAL SERIAL IMPORTS --------
try:
    import serial
    import serial.tools.list_ports as list_ports
except Exception:
    serial = None
    list_ports = None

    # PR05 THEMES ##MARK## THEME — Monokai QSS (dark)
# --- THEME: Monokai (dark) ---
MONOKAI_DARK_QSS = r"""
QMainWindow, QWidget { background:#1e1f1c; color:#f8f8f2; }
QTabBar::tab{background:#2a2b26;color:#f8f8f2;padding:7px 12px;margin:2px;border-top-left-radius:8px;border-top-right-radius:8px;}
QTabBar::tab:selected{background:#49483e;color:#a6e22e;}
QTabWidget::pane{border:1px solid #3a3b36;top:-1px;}

QGroupBox{border:1px solid #3a3b36;border-radius:8px;margin-top:12px;}
QGroupBox::title{left:8px;padding:0 4px;color:#a6e22e;}

QLineEdit,QSpinBox,QDoubleSpinBox,QComboBox,QTextEdit,QPlainTextEdit{
  background:#272822;color:#f8f8f2;border:1px solid #3a3b36;border-radius:6px;padding:4px;
}
QPushButton{background:#3a3b36;color:#f8f8f2;border:1px solid #565751;border-radius:6px;padding:6px 10px;}
QPushButton:hover{background:#49483e;}
QPushButton:disabled{background:#2a2b26;color:#8f908a;}

/* ============== SPINBOX ARROWS (dark) ============== */
QSpinBox::up-button,QDoubleSpinBox::up-button{
  subcontrol-origin:border; subcontrol-position:top right; width:18px;
  background:#3a3b36; border-left:1px solid #565751; border-top-right-radius:6px;
}
QSpinBox::down-button,QDoubleSpinBox::down-button{
  subcontrol-origin:border; subcontrol-position:bottom right; width:18px;
  background:#3a3b36; border-left:1px solid #565751; border-bottom-right-radius:6px;
}

"""

# --- THEME: Your creamy light variant (high-contrast arrows) ---
MONOKAI_CREAM_QSS = r"""
QMainWindow, QWidget { background:#f8f8f2; color:#20201e; }
QTabBar::tab{background:#e4e1b7;color:#2e4401;border-top-left-radius:8px;border-top-right-radius:8px;padding:8px;margin:2px;}
QTabBar::tab:selected{background:#A6E22E;color:#272822;}
QTabWidget::pane{border:1px solid #506b39; top:-1px;}

QGroupBox{border:2px solid #506b39;border-radius:12px;margin-top:12px;}
QGroupBox::title{left:10px;padding:0 6px;color:#2f4600;}

QLineEdit,QSpinBox,QDoubleSpinBox,QComboBox,QTextEdit,QPlainTextEdit{
  background:#f8f8f0;color:#18390c;border:2px solid #506b39;border-radius:8px;padding:4px;
}
QPushButton{background:#A6E22E;color:#272822;border:2px solid #506b39;border-radius:8px;padding:6px 10px;}
QPushButton:hover{background:#C0FF3E;}

/* Table bits kept from your snippet omitted for brevity */

/* ============== SPINBOX ARROWS (light, no images) ============== */
QSpinBox::up-button, QDoubleSpinBox::up-button {
  subcontrol-origin: border;
  subcontrol-position: top right;
  width: 22px;                      /* make the cap visible */
  background: #bbcc92;              /* contrast cap */
  border-left: 2px solid #506b39;
  border-top-right-radius: 8px;
}
QSpinBox::down-button, QDoubleSpinBox::down-button {
  subcontrol-origin: border;
  subcontrol-position: bottom right;
  width: 22px;
  background: #e2afa1;
  border-left: 2px solid #506b39;
  border-bottom-right-radius: 8px;
}
QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover{
  background: #c6d78f;              /* hover feedback */
}
QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover{ 
background:#f9d6cc;
} 

QSpinBox::up-button:pressed, QDoubleSpinBox::up-button:pressed
{
  background: #b6c97f;              /* pressed feedback */
}
QSpinBox::down-button:pressed, QDoubleSpinBox::down-button:pressed { 
background:#c95852; 
}
"""


def apply_theme(window, variant="dark"):
    QApplication.setStyle(QStyleFactory.create("Fusion"))

    pal = window.palette()
    if variant == "light":
        window.setStyleSheet(MONOKAI_CREAM_QSS)
        pal.setColor(QPalette.Window, QColor("#f8f8f2"))
        pal.setColor(QPalette.Base,   QColor("#f8f8f0"))
        pal.setColor(QPalette.Text,   QColor("#20201e"))
        pal.setColor(QPalette.Button, QColor("#d0e3a1"))
        pal.setColor(QPalette.ButtonText, QColor("#2f4600"))
        pal.setColor(QPalette.Highlight, QColor("#A6E22E"))
        pal.setColor(QPalette.HighlightedText, QColor("#272822"))
    else:
        window.setStyleSheet(MONOKAI_DARK_QSS)
        pal.setColor(QPalette.Window, QColor("#1e1f1c"))
        pal.setColor(QPalette.Base,   QColor("#272822"))
        pal.setColor(QPalette.Text,   QColor("#f8f8f2"))
        pal.setColor(QPalette.Button, QColor("#3a3b36"))
        pal.setColor(QPalette.ButtonText, QColor("#f8f8f2"))
        pal.setColor(QPalette.Highlight, QColor("#49483e"))
        pal.setColor(QPalette.HighlightedText, QColor("#a6e22e"))

    window.setPalette(pal)

    # keep custom-painted widgets in sync
    if hasattr(window, "plot") and hasattr(window.plot, "apply_theme_from_palette"):
        window.plot.apply_theme_from_palette(pal)
    if hasattr(window, "endLong"):  window.endLong.update()
    if hasattr(window, "endShort"): window.endShort.update()

# -------- Parser (module scope) --------
# Accepts lines like: P1X1000Y2000R3000S4000  P2X...  (order doesn't matter per P)
P_BLOCK = re.compile(
    r"P(?P<p>[1-4])\s*X(?P<X>\d+)\s*Y(?P<Y>\d+)\s*R(?P<R>\d+)\s*S(?P<S>\d+)",
    re.IGNORECASE,
)

def parse_p_line(line: str) -> dict:
    # strip Arduino monitor prefix like "12:34:56.789 -> "
    for sep in ("->", "=>"):
        if sep in line:
            line = line.split(sep, 1)[1]
            break
    out = {}
    for m in P_BLOCK.finditer(line):
        p = int(m.group("p"))
        base = f"p{p}"
        out[f"{base}X"] = int(m.group("X"))
        out[f"{base}Y"] = int(m.group("Y"))
        out[f"{base}R"] = int(m.group("R"))
        out[f"{base}S"] = int(m.group("S"))
    return out

# -------- End-view Widget --------
class EndViewWidget(QWidget):
    changed = pyqtSignal(int)
    def __init__(self, side_name="Long", parent=None):
        super().__init__(parent)
        self.setMinimumSize(QSize(160, 160))
        self.side_name = side_name
        self.selected = 0
    def setSelected(self, idx: int):
        self.selected = int(max(0, min(2, idx)))
        self.update(); self.changed.emit(self.selected)
    def mousePressEvent(self, e):
        idx = self._hitTest(e.pos().x(), e.pos().y())
        if idx is not None:
            self.setSelected(idx)
    def _pouch_centers(self, R):
        import math
        cx, cy = self.width()/2, self.height()/2
        r = R * 0.55
        angles = [150, -90, 30]
        pts = []
        for ang in angles:
            rad = math.radians(ang)
            px = cx + r * math.cos(rad)
            py = cy + r * math.sin(rad)
            pts.append((px, py))
        return pts
    def _hitTest(self, x, y):
        R = min(self.width(), self.height()) * 0.45
        pts = self._pouch_centers(R)
        for i, (px, py) in enumerate(pts):
            if (x - px) ** 2 + (y - py) ** 2 <= (R*0.12) ** 2:
                return i
        return None
    def paintEvent(self, _):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing, True)
        W, H = self.width(), self.height()
        cx, cy = W/2, H/2
        R = min(W, H) * 0.45
        p.fillRect(0, 0, W, H, QColor(18, 18, 18))
        pen = QPen(QColor(220, 220, 220)); pen.setWidth(1); p.setPen(pen)
        p.drawText(8, 18, f"Section: {self.side_name}")
        p.setPen(QPen(QColor(200, 200, 200), 2)); p.setBrush(Qt.NoBrush)
        p.drawEllipse(int(cx - R), int(cy - R), int(2*R), int(2*R))
        pts = self._pouch_centers(R)
        for i, (px, py) in enumerate(pts):
            color = QColor(90, 90, 90)
            if i == self.selected:
                color = QColor(0, 180, 255) if self.side_name == "Long" else QColor(255, 120, 0)
            p.setBrush(QBrush(color))
            p.setPen(QPen(QColor(230, 230, 230), 1))
            d = R * 0.24
            p.drawEllipse(int(px - d/2), int(py - d/2), int(d), int(d))
            p.setPen(QPen(QColor(255, 255, 255), 1))
            p.drawText(int(px - 4), int(py + 4), str(i+1))

# -------- Serial Worker --------
class SerialWorker(QObject):
    line_batch = pyqtSignal(list)   # list[str]
    status = pyqtSignal(str)
    connected = pyqtSignal(bool)

    def __init__(self):
        super().__init__()
        self._ser = None
        self._stop = threading.Event()
        self._buf = bytearray()
        self._rx_thread = None
        self._txq: "queue.Queue[bytes]" = queue.Queue(maxsize=1000)
        self._last_err = ""
        self._err_silence_until = 0.0
        # runtime tx/rx state
        self._opened_at = 0.0
        self._last_rx_time = 0.0
        self._ready = False
        self._awaiting_ack = False
        self._ack_regex = None
        self._line_ending = b"\r\n"
        self._tx_min_interval = 0.10
        self._last_tx_time = 0.0
        self._wr_timeouts = 0
        self._tx_block_until = 0.0
        self._warmup_ms = 800

    # ----- configuration from GUI -----
    def set_line_ending(self, eol: str):
        if eol == "LF (\\n)": self._line_ending = b"\n"
        elif eol == "CRLF (\\r\\n)": self._line_ending = b"\r\n"
        elif eol == "CR (\\r)": self._line_ending = b"\r"
    def set_ack_regex(self, s: str):
        try:
            self._ack_regex = re.compile(s) if s.strip() else None
        except re.error:
            self._ack_regex = None
    def set_warmup_ms(self, ms: int):
        self._warmup_ms = max(0, int(ms))

    def list_ports(self):
        if list_ports is None: return []
        return [p.device for p in list_ports.comports()]

    def connect(self, port, baud=9600):
        if serial is None:
            self.status.emit("pyserial not installed. Run: pip install pyserial")
            return False
        try:
            self._ser = serial.Serial(
                port,
                baudrate=baud,
                timeout=0.05,
                write_timeout=1.00,
                rtscts=False,
                dsrdtr=False,
                exclusive=True,
            )
            # Arduino/ESP style open
            try:
                self._ser.setRTS(False)
                self._ser.setDTR(False)
                time.sleep(0.05)
                self._ser.setDTR(True)
            except Exception:
                pass
            try:
                self._ser.reset_input_buffer()
                self._ser.reset_output_buffer()
            except Exception:
                pass

            self._stop.clear()
            self._opened_at = time.monotonic()
            self._last_rx_time = 0.0
            self._ready = False
            self._awaiting_ack = False
            self._wr_timeouts = 0
            self._tx_block_until = 0.0

            self._rx_thread = threading.Thread(target=self._io_loop, daemon=True)
            self._rx_thread.start()
            self.connected.emit(True)
            self.status.emit(f"Connected to {port} @ {baud}")
            return True
        except Exception as e:
            self._ser = None
            self.connected.emit(False)
            self.status.emit(f"Connect failed: {e}")
            return False

    def disconnect(self):
        self._stop.set()
        if self._ser:
            try: self._ser.close()
            except Exception: pass
        self._ser = None
        self.connected.emit(False)
        self.status.emit("Disconnected")

    # Public TX from GUI:
    def send_json(self, obj):
        self._enqueue_tx((json.dumps(obj)).encode("utf-8") + self._line_ending)
    def send_text(self, s: str):
        self._enqueue_tx(s.encode("utf-8") + self._line_ending)

    def resume_tx_now(self):
        # manual unpause
        self._tx_block_until = 0.0
        self._wr_timeouts = 0
        self._ready = True  # force-ready; useful if device doesn't echo
        self.status.emit("TX resumed manually")

    def _enqueue_tx(self, b: bytes):
        try:
            self._txq.put_nowait(b)
        except queue.Full:
            # drop oldest to keep latest
            try: self._txq.get_nowait()
            except queue.Empty: pass
            self._txq.put_nowait(b)

    def _emit_err(self, msg: str):
        now = time.monotonic()
        if msg == self._last_err and now < self._err_silence_until:
            return
        self._last_err = msg
        self._err_silence_until = now + 0.5
        self.status.emit(msg)

    def _io_loop(self):
        pending = []
        next_emit = time.monotonic() + 0.03   # 30 ms batch cadence
        backoff = 0.0
        while not self._stop.is_set():
            try:
                ser = self._ser
                if ser is None:
                    break

                # ---- TX first (input-first), with gating ----
                now = time.monotonic()
                warmup_ok = (now - self._opened_at) * 1000.0 >= self._warmup_ms
                can_tx = (self._ready or warmup_ok) and (now >= self._tx_block_until)
                if can_tx and (now - self._last_tx_time) >= self._tx_min_interval:
                    try:
                        tx = self._txq.get_nowait()
                        try:
                            ser.write(tx)
                            ser.flush()
                            self._last_tx_time = now
                            self._wr_timeouts = 0
                            if self._ack_regex is not None:
                                self._awaiting_ack = True
                        except serial.SerialTimeoutException:
                            self._wr_timeouts += 1
                            if self._wr_timeouts >= 3:
                                self._tx_block_until = now + 2.0
                                self._drain_tx_backlog(max_keep=2)
                                self._emit_err("Write timeout: pausing TX for 2s")
                            else:
                                self._tx_block_until = now + 0.20
                        except Exception as e:
                            self._emit_err(f"Write error: {e}")
                    except queue.Empty:
                        pass

                # ---- RX bulk ----
                chunk = ser.read(4096)
                if chunk:
                    self._ready = True
                    self._last_rx_time = time.monotonic()
                    backoff = 0.0
                    self._buf.extend(chunk)
                    while True:
                        i = self._buf.find(b'\n')
                        if i < 0: break
                        line = self._buf[:i].rstrip(b'\r').decode('utf-8', 'replace')
                        del self._buf[:i+1]
                        # ACK gating
                        if self._ack_regex is not None and self._awaiting_ack:
                            if self._ack_regex.search(line):
                                self._awaiting_ack = False
                        pending.append(line)

                now = time.monotonic()
                if pending and now >= next_emit:
                    if len(pending) > 1000:
                        pending = pending[-1000:]
                    batch = pending; pending = []
                    self.line_batch.emit(batch)
                    next_emit = now + 0.03

                if not chunk:
                    time.sleep(0.002 + backoff)

            except Exception as e:
                self._emit_err(f"Read error: {e}")
                backoff = min(0.250, (0.005 if backoff == 0 else backoff * 2))
                time.sleep(backoff)

    def _drain_tx_backlog(self, max_keep=0):
        """Drop queued commands to recover; keep the last N."""
        kept = []
        try:
            while True:
                kept.append(self._txq.get_nowait())
        except queue.Empty:
            pass
        if max_keep > 0:
            kept = kept[-max_keep:]
        else:
            kept = []
        for b in kept:
            try: self._txq.put_nowait(b)
            except queue.Full: break

# -------- Lightweight realtime plot --------
class RealtimePlot(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(600, 320)
        self.window_ms = 10_000
        self.series: dict[str, deque[tuple[float, float]]] = {}
        self.visible: set[str] = set()
        self.maxlen = 50_000
        self._bg = QColor(10, 10, 10)
        self._grid_on = True
        self._auto_y = True
        self._ymin = 0.0
        self._ymax = 100.0
        self._x_tick_ms = None
        self._palette = [
            QColor(0, 200, 255), QColor(255, 140, 0),
            QColor(120, 200, 80), QColor(220, 120, 220),
            QColor(250, 220, 80), QColor(120, 160, 255),
        ]
    def set_window_ms(self, ms: int): self.window_ms = max(25, int(ms)); self.update()
    def set_auto_y(self, on: bool): self._auto_y = bool(on); self.update()
    def set_grid(self, on: bool): self._grid_on = bool(on); self.update()
    def set_x_tick_ms(self, step_ms): self._x_tick_ms = step_ms; self.update()
    def clear(self): self.series.clear(); self.visible.clear(); self.update()
    def push(self, name: str, t_ms: float, y: float):
        dq = self.series.get(name)
        if dq is None:
            dq = self.series[name] = deque(maxlen=self.maxlen)
            self.visible.add(name)
        dq.append((t_ms, y))
        self.update()
    def add_sample(self, name, t_ms, y): self.push(name, t_ms, y)
    def set_visible(self, name: str, on: bool):
        if on: self.visible.add(name)
        else: self.visible.discard(name)
        self.update()
    def _nice_step(self, span):
        if span <= 0: return 1000
        raw = span / 8.0
        base = 10 ** int(len(str(int(max(1, raw)))) - 1)
        for m in (1, 2, 5, 10):
            step = m * base
            if raw <= step: return step
        return 10 * base
    def _auto_scale_y(self, tmin):
        if not self._auto_y: return
        vals = []
        for name in self.visible:
            for (t, y) in self.series.get(name, ()):
                if t >= tmin: vals.append(y)
        if not vals: return
        lo, hi = min(vals), max(vals)
        if hi == lo: hi = lo + 1.0
        pad = 0.1 * (hi - lo)
        self._ymin, self._ymax = lo - pad, hi + pad
    def paintEvent(self, _):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing, True)
        W, H = self.width(), self.height()
        p.fillRect(0, 0, W, H, self._bg)
        margin_l, margin_r, margin_t, margin_b = 56, 16, 20, 34
        x0, y0 = margin_l, H - margin_b
        x1, y1 = W - margin_r, margin_t
        plot_w = max(1, x1 - x0); plot_h = max(1, y0 - y1)
        p.setPen(QPen(QColor(180, 180, 180), 1)); p.drawRect(x0, y1, plot_w, plot_h)
        total_pts = sum(len(dq) for dq in self.series.values())
        if total_pts < 2:
            p.setPen(QPen(QColor(200, 200, 200))); p.drawText(x0 + 8, y1 + 20, "Waiting for serial data..."); return
        tmax = max((dq[-1][0] for dq in self.series.values() if dq), default=0.0)
        tmin = tmax - self.window_ms
        self._auto_scale_y(tmin)
        ymin, ymax = self._ymin, self._ymax
        def map_x(t):
            if t <= tmin: return x0
            if t >= tmax: return x1
            return x0 + int((t - tmin) / max(1e-9, (tmax - tmin)) * plot_w)
        def map_y(y):
            if y <= ymin: return y0
            if y >= ymax: return y1
            return y0 - int((y - ymin) / max(1e-9, (ymax - ymin)) * plot_h)
        if self._grid_on:
            vpen = QPen(QColor(60, 60, 60), 1, Qt.DashLine)
            p.setPen(vpen)
            xstep = self._x_tick_ms if self._x_tick_ms else self._nice_step(self.window_ms)
            first_tick = int((tmin // xstep) * xstep)
            tick = first_tick
            while tick <= tmax:
                X = map_x(tick); p.drawLine(X, y1, X, y0); tick += xstep
            p.setPen(vpen)
            ticks = 6
            for i in range(ticks + 1):
                yy = ymin + (ymax - ymin) * i / ticks
                Y = map_y(yy); p.drawLine(x0, Y, x1, Y)
                p.setPen(QPen(QColor(170, 170, 170))); p.drawText(x0 - 48, Y + 4, f"{yy:.1f}")
                p.setPen(vpen)
        for idx, name in enumerate([n for n in self.series if n in self.visible]):
            color = self._palette[idx % len(self._palette)]
            pen = QPen(color, 2); p.setPen(pen)
            prev = None
            for t, y in self.series[name]:
                if t < tmin: continue
                X, Y = map_x(t), map_y(y)
                if prev is not None: p.drawLine(prev[0], prev[1], X, Y)
                prev = (X, Y)
        p.setPen(QPen(QColor(220, 220, 220)))
        lx, ly = x0 + 8, y1 - 8
        names = [n for n in self.series if n in self.visible]
        for idx, name in enumerate(names[:8]):
            color = self._palette[idx % len(self._palette)]
            p.setBrush(QBrush(color)); p.setPen(Qt.NoPen); p.drawRect(lx, ly - 12, 12, 12)
            p.setPen(QPen(QColor(220, 220, 220))); p.drawText(lx + 16, ly - 2, name)
            lx += 80

# -------- CSV Writer thread (full-rate logging) --------
class CSVWriter(threading.Thread):
    HEADER = ["sample", "t_ms",
              "p1X","p1Y","p1R","p1S",
              "p2X","p2Y","p2R","p2S",
              "p3X","p3Y","p3R","p3S",
              "p4X","p4Y","p4R","p4S"]
    def __init__(self, base_dir: Path):
        super().__init__(daemon=True)
        self.base_dir = Path(base_dir)
        d = time.strftime("%Y%m%d")
        log_dir = self.base_dir / "logs" / d
        log_dir.mkdir(parents=True, exist_ok=True)
        self._path = log_dir / time.strftime("serial_%Y%m%d_%H%M%S.csv")
        self._fp = open(self._path, "w", newline="")
        self._w = csv.writer(self._fp)
        self._w.writerow(self.HEADER); self._fp.flush()
        self.q: "queue.Queue[list | None]" = queue.Queue(maxsize=5000)
        self.running = True
    @property
    def path(self): return str(self._path)
    def run(self):
        n = 0
        while self.running:
            item = self.q.get()
            if item is None: break
            self._w.writerow(item); n += 1
            if (n % 100) == 0:
                self._fp.flush()
        self._fp.flush(); self._fp.close()
    def close(self):
        self.running = False
        try: self.q.put_nowait(None)
        except Exception: pass

# -------- Main Window --------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Vine Robot Controller — ADC Setpoint — Tabs + Auto CSV + Multi-Series Plot")
        self.setMinimumSize(1200, 640)
        self._t0 = time.monotonic()
        self.sample = 0

        # Serial worker & CSV
        self.ser = SerialWorker()
        self.ser.line_batch.connect(self.on_serial_batch)
        self.ser.status.connect(self.on_status)
        self.ser.connected.connect(self.on_connected)

        self._base_dir = Path(__file__).resolve().parent
        self.csv = CSVWriter(self._base_dir); self.csv.start()

        # ===== UI =====
        tabs = QTabWidget(); self.setCentralWidget(tabs)

        # Tab 1: Operate
        tab_operate = QWidget(); tabs.addTab(tab_operate, "Operate")
        op_root = QHBoxLayout(tab_operate)

        # Left: end-views
        leftCol = QVBoxLayout()
        self.endLong = EndViewWidget("Long"); self.endShort = EndViewWidget("Short")
        evRow = QHBoxLayout(); evRow.addWidget(self.endLong); evRow.addWidget(self.endShort)
        leftCol.addLayout(evRow)
        statusRow = QHBoxLayout()
        
        # PR04 ##MARK## NEW ── Motor control bar (reverse← 0 →forward)
        motorBox = QGroupBox("Motor control")
        hb = QHBoxLayout(motorBox)

        self.sldMotor = QSlider(Qt.Horizontal)
        self.sldMotor.setRange(-255, 255)          # - = REV, + = FWD
        self.sldMotor.setSingleStep(5)
        self.sldMotor.setPageStep(20)
        self.sldMotor.setValue(0)
        self.sldMotor.setMinimumWidth(260)
        # **style: red→gray→green**
        self.sldMotor.setStyleSheet("""
        QSlider::groove:horizontal {
          height: 10px; border-radius: 5px;
          background: qlineargradient(x1:0, y1:0.5, x2:1, y2:0.5,
            stop:0   #d84a4a, stop:0.495 #5a5a5a,
            stop:0.505 #5a5a5a, stop:1   #3ac96b);
        }
        QSlider::handle:horizontal {
          background: #ffffff; border:1px solid #888;
          width:16px; margin:-4px 0; border-radius:8px;
        }
        """)

        self.spnMotor = QSpinBox()
        self.spnMotor.setRange(-255, 255)
        self.spnMotor.setValue(0)
        self.spnMotor.setFixedWidth(70)

        self.btnMotorSend = QPushButton("Send Motor")
        self.btnMotorStop = QPushButton("Stop")

        # wire slider<->spin (no sending yet)
        self.sldMotor.valueChanged.connect(self.spnMotor.setValue)
        self.spnMotor.valueChanged.connect(self.sldMotor.setValue)

        # place widgets
        hb.addWidget(QLabel("REV"))
        hb.addWidget(self.sldMotor, 1)
        hb.addWidget(QLabel("FWD"))
        hb.addSpacing(8)
        hb.addWidget(self.spnMotor)
        hb.addSpacing(6)
        hb.addWidget(self.btnMotorSend)
        hb.addWidget(self.btnMotorStop)

        leftCol.addWidget(motorBox)   # put where the old labels were

        op_root.addLayout(leftCol, 3)

        # Middle: controls
        ctrlBox = QGroupBox("Controls"); form = QFormLayout(ctrlBox)
        self.cmbPort = QComboBox(); self.cmbPort.addItems(self.ser.list_ports())
        self.btnRefresh = QPushButton("Refresh"); self.btnConn = QPushButton("Connect")
        self.btnDisc = QPushButton("Disconnect"); self.btnDisc.setEnabled(False)
        portRow = QHBoxLayout()
        portRow.addWidget(self.cmbPort); portRow.addWidget(self.btnRefresh)
        portRow.addWidget(self.btnConn); portRow.addWidget(self.btnDisc)
        form.addRow(QLabel("Serial Port:"), portRow)

        # EOL & ACK
        self.cmbEOL = QComboBox(); self.cmbEOL.addItems(["CRLF (\\r\\n)", "LF (\\n)", "CR (\\r)"]); self.cmbEOL.setCurrentIndex(0)
        self.txtAck = QLineEdit("^(OK|ACK)$")
        self.spnWarm = QSpinBox(); self.spnWarm.setRange(0, 5000); self.spnWarm.setValue(800)
        eolRow = QHBoxLayout()
        eolRow.addWidget(QLabel("EOL:")); eolRow.addWidget(self.cmbEOL)
        eolRow.addWidget(QLabel("ACK regex:")); eolRow.addWidget(self.txtAck)
        eolRow.addWidget(QLabel("Warm-up (ms):")); eolRow.addWidget(self.spnWarm)
        form.addRow(eolRow)

        self.spinPSI = QDoubleSpinBox(); self.spinPSI.setRange(0.0, 4000.0)
        self.spinPSI.setSingleStep(25); self.spinPSI.setValue(25)
        form.addRow(QLabel("Target PSI:"), self.spinPSI)
        self.btnSetLong = QPushButton("Send SET (Long)")
        self.btnSetShort = QPushButton("Send SET (Short)")
        self.btnHoldLong = QPushButton("HOLD (Long)")
        self.btnHoldShort = QPushButton("HOLD (Short)")
        
        # PR03 **TODO [UI] — add exhaust duration field (optional)**
        self.spinExhMs = QSpinBox()
        self.spinExhMs.setRange(50, 5000)
        self.spinExhMs.setSingleStep(50)
        self.spinExhMs.setValue(800)
        self.spinExhMs.setSuffix(" ms")
        form.addRow(QLabel("Exhaust ms:"), self.spinExhMs)

        self.btnStop = QPushButton("STOP (All)")
        self.btnResumeTX = QPushButton("Resume TX")
        form.addRow(self.btnSetLong, self.btnSetShort)
        form.addRow(self.btnHoldLong, self.btnHoldShort)
        # PR03 **TODO [UI] — add EXHAUST buttons**
        self.btnExhLong  = QPushButton("EXHAUST (Long)")
        self.btnExhShort = QPushButton("EXHAUST (Short)")
        form.addRow(self.btnExhLong, self.btnExhShort)

        form.addRow(self.btnStop, self.btnResumeTX)

        self.chkEcho = QCheckBox("Echo incoming to Quick Log"); self.chkEcho.setChecked(False)
        form.addRow(self.chkEcho)
        self.chkLivePlot = QCheckBox("Live plot"); self.chkLivePlot.setChecked(True)
        form.addRow(self.chkLivePlot)
        self.chkCSV = QCheckBox("Write CSV"); self.chkCSV.setChecked(True)
        form.addRow(self.chkCSV)
        self.lblCsvPath = QLabel(f"Logging to: {self.csv.path}")
        form.addRow(self.lblCsvPath)
        op_root.addWidget(ctrlBox, 2)

        # Right: Quick Log
        logBox = QGroupBox("Quick Log"); logLay = QVBoxLayout(logBox)
        self.txtLog = QPlainTextEdit(); self.txtLog.setReadOnly(True)
        logLay.addWidget(self.txtLog); op_root.addWidget(logBox, 3)

        # Wire buttons
        self.btnRefresh.clicked.connect(self.on_refresh)
        self.btnConn.clicked.connect(self.on_connect)
        self.btnDisc.clicked.connect(self.on_disconnect)
        self.btnSetLong.clicked.connect(lambda: self.send_set("Long"))
        self.btnSetShort.clicked.connect(lambda: self.send_set("Short"))
        self.btnHoldLong.clicked.connect(lambda: self.send_hold("Long"))
        self.btnHoldShort.clicked.connect(lambda: self.send_hold("Short"))
        
        # PR03 **TODO [signals]**
        self.btnExhLong.clicked.connect(lambda: self.send_exhaust("Long"))
        self.btnExhShort.clicked.connect(lambda: self.send_exhaust("Short"))

        self.btnStop.clicked.connect(self.send_stop)
        self.btnResumeTX.clicked.connect(self.on_resume_tx)

        # PR04 signals connect the buttons (put with your other connects)
        self.btnMotorSend.clicked.connect(self.on_send_motor)
        self.btnMotorStop.clicked.connect(self.on_stop_motor)

        # Tab 2: Plot
        tab_plot = QWidget(); tabs.addTab(tab_plot, "Plot")
        pl_root = QVBoxLayout(tab_plot)
        topRow = QHBoxLayout()
        self.plot = RealtimePlot()
        self.cmbWindow = QComboBox(); self.cmbWindow.addItems(
            ["1000 ms","2000 ms","5000 ms","10000 ms","20000 ms","60000 ms"]
        ); self.cmbWindow.setCurrentText("10000 ms")
        self.cmbTick = QComboBox(); self.cmbTick.addItems(
            ["auto","25 ms","50 ms","1000 ms","2000 ms","5000 ms"]
        ); self.cmbTick.setCurrentText("auto")
                # PR02 ** TODO ADD METRIC in Plot tab UI (near your current P1Y..P4Y checkboxes) ---
                # replace the Metric QComboBox with checkboxes:
        self.chkMX = QCheckBox("X"); self.chkMY = QCheckBox("Y")
        self.chkMR = QCheckBox("R"); self.chkMS = QCheckBox("S")
        self.chkMY.setChecked(True)       # sensible default: show Y only

        topRow.addWidget(QLabel("Metric"))
        for m in (self.chkMX, self.chkMY, self.chkMR, self.chkMS):
            topRow.addWidget(m)

        # keep your pouch toggles as P1..P4 (not P1Y etc.):
        self.chkP1 = QCheckBox("P1"); self.chkP1.setChecked(True)
        self.chkP2 = QCheckBox("P2"); self.chkP2.setChecked(True)
        self.chkP3 = QCheckBox("P3"); self.chkP3.setChecked(True)
        self.chkP4 = QCheckBox("P4"); self.chkP4.setChecked(True)
        topRow.addSpacing(12)
        for w in (self.chkP1, self.chkP2, self.chkP3, self.chkP4):
            topRow.addWidget(w)

        topRow.addWidget(QLabel("Window")); topRow.addWidget(self.cmbWindow)
        topRow.addSpacing(12); topRow.addWidget(QLabel("X grid step")); topRow.addWidget(self.cmbTick)
        topRow.addStretch(1)
        for w in (self.chkP1, self.chkP2, self.chkP3, self.chkP4): topRow.addWidget(w)
        pl_root.addLayout(topRow); pl_root.addWidget(self.plot, 1)
        self.cmbWindow.currentTextChanged.connect(self.on_window_change)
        self.cmbTick.currentTextChanged.connect(self.on_tick_change)
        self.repaint_timer = QTimer(self); self.repaint_timer.timeout.connect(self.plot.update); self.repaint_timer.start(100)

        # Tab 3: Raw Log
        tab_log = QWidget(); tabs.addTab(tab_log, "Log")
        tlog = QVBoxLayout(tab_log)
        self.txtLog2 = QPlainTextEdit(); self.txtLog2.setReadOnly(True)
        tlog.addWidget(self.txtLog2)

        # Buffered UI log
        self._log_buffer = []
        self._log_timer = QTimer(self); self._log_timer.timeout.connect(self._flush_log); self._log_timer.start(100)
        self._append_startup_log()

        # Plot decimation
        self._last_plot_emit_ms = 0.0
        self._plot_min_interval_ms = 15.0

        # UI snapshot timer
        self._latest_snapshot = None
        self._ui_timer = QTimer(self); self._ui_timer.timeout.connect(self._drain_snapshot); self._ui_timer.start(100)

        # PR05 THEMES
        apply_theme(self,'light')
        # **OPTIONAL** menu action to toggle on/off
        view = self.menuBar().addMenu("View")
        actDark  = view.addAction("Dark Monokai");  actDark.triggered.connect(lambda: apply_theme(self,"dark"))
        actLight = view.addAction("Creamy Light");  actLight.triggered.connect(lambda: apply_theme(self,"light"))
        self.actTheme = QAction("Toggle Monokai", self, checkable=True, checked=True)
        self.actTheme.toggled.connect(lambda on: apply_theme(self) if on else self.setStyleSheet(""))
        view.addAction(self.actTheme)


    # ----- UI helpers -----
    def _append_startup_log(self):
        self.append_log("Tip: SET/HOLD/STOP send simple JSON or text to your device.")
        self.append_log(f"Auto CSV logging to: {self.csv.path}")
        self.append_log("Display is throttled; CSV stores full-rate data.")

    def append_log(self, s: str):
        self._log_buffer.append(str(s))

    def _flush_log(self):
        if not self._log_buffer: return
        MAX_LINES = 5000
        if len(self._log_buffer) > MAX_LINES:
            self._log_buffer = self._log_buffer[-MAX_LINES:]
        out = "\n".join(self._log_buffer[:400]); del self._log_buffer[:400]
        self.txtLog.appendPlainText(out)
        self.txtLog.moveCursor(self.txtLog.textCursor().End)

    # ----- Controls -----
    def on_refresh(self):
        self.cmbPort.clear(); self.cmbPort.addItems(self.ser.list_ports())

    def on_connect(self):
        # push EOL / ACK / warmup settings before open
        self.ser.set_line_ending(self.cmbEOL.currentText())
        self.ser.set_ack_regex(self.txtAck.text())
        self.ser.set_warmup_ms(self.spnWarm.value())

        port = self.cmbPort.currentText()
        if not port:
            QMessageBox.warning(self, "No port", "Select a serial port"); return
        if self.ser.connect(port, 9600):
            self.btnConn.setEnabled(False); self.btnDisc.setEnabled(True)
            try: self.csv.close()
            except Exception: pass
            self.csv = CSVWriter(self._base_dir); self.csv.start()
            self.lblCsvPath.setText(f"Logging to: {self.csv.path}")
            self.append_log(f"New CSV: {self.csv.path}")

    def on_disconnect(self):
        self.ser.disconnect()
        self.btnConn.setEnabled(True); self.btnDisc.setEnabled(False)

    def on_connected(self, ok):
        self.append_log("CONNECTED" if ok else "DISCONNECTED")

    def on_status(self, msg):
        self.append_log(msg)

    def on_resume_tx(self):
        self.ser.resume_tx_now()

    # ----- Timescale -----
    def on_window_change(self, txt):
        ms = int(txt.split()[0]); self.plot.set_window_ms(ms)
    def on_tick_change(self, txt):
        self.plot.set_x_tick_ms(None if txt == "auto" else int(txt.split()[0]))

    # ----- Commands -----
    def current_seg(self, side):
        return self.endLong.selected if side == "Long" else self.endShort.selected

    def _debounce(self, btn: QPushButton, ms=150):
        btn.setEnabled(False); QTimer.singleShot(ms, lambda: btn.setEnabled(True))

    def send_set(self, side):
        seg = int(self.current_seg(side)) + 1
        if side == "Long": seg += 3
        pouch = f"p{seg}"; value = int(round(self.spinPSI.value()))
        msg = f"{pouch} x{value}"
        self.ser.send_text(msg)
        self.append_log(f">> {msg}")
        self._debounce(self.btnSetLong if side=="Long" else self.btnSetShort)

    def send_hold(self, side):
        seg = int(self.current_seg(side))
        payload = {"cmd": "HOLD", "section": side, "seg": seg}
        self.ser.send_json(payload); self.append_log(f">> {payload}")
        self._debounce(self.btnHoldLong if side=="Long" else self.btnHoldShort)

    # PR03 ##MARK## NEW — send_exhaust: like Send SET, but EXH
    def send_exhaust(self, side: str):
        seg = int(self.current_seg(side)) + 1     # 1..3 within a side
        if side == "Long":
            seg += 3                               # 4..6 overall
        ms = int(self.spinExhMs.value()) if hasattr(self, "spinExhMs") else 800

        # &&NOTE&& Choose one protocol. Default below is a self-describing TEXT line.
        # If your firmware already expects another legacy text format, flip PROTOCOL.

        PROTOCOL = "PSET"   # options: "PSET" (recommended) or "LEGACY" (old style)

        if PROTOCOL == "PSET":
            # Human-readable; easy to debug in Serial Monitor
            cmd = f"PSET N={seg} MODE=EXH DURATION_MS={ms}"
            self.ser.send_text(cmd)
            self.append_log(f">> {cmd}")
        elif PROTOCOL == "LEGACY":
            # Matches your old Send SET style (text). Adjust if your MCU expects different token.
            # Example: " p{seg} e{ms}"
            msg = f" p{seg} e{ms}"
            self.ser.send_text(msg)
            self.append_log(f">> {msg}")
        else:
            # JSON alternative (if your firmware parses JSON)
            payload = {"cmd": "PSET", "n": seg, "mode": "EXH", "ms": ms}
            self.ser.send_json(payload)
            self.append_log(f">> {payload}")
    
       # PR04 ##MARK## NEW ── one-shot send based on current slider value
    def on_send_motor(self):
        val = int(self.sldMotor.value())
        # deadband (optional for small jitter)
        if -3 <= val <= 3:
            val = 0
        direction = "FWD" if val >= 0 else "REV"
        pwm = abs(val)
        # **text protocol** (simple + debuggable). Change if your firmware expects JSON.
        cmd = f"MOTOR DIR={direction} PWM={pwm}"
        self.ser.send_text(cmd)
        self.append_log(f">> {cmd}")

    def on_stop_motor(self):
        # normalize stop as FWD 0 (or send a dedicated BRAKE/HOLD if your firmware has it)
        cmd = "MOTOR DIR=FWD PWM=0"
        self.ser.send_text(cmd)
        self.append_log(f">> {cmd}")
        self.sldMotor.setValue(0)  # snap UI back to center

    def send_stop(self):
        payload = {"cmd": "STOP"}
        self.ser.send_json(payload); self.append_log(f">> {payload}")
        self._debounce(self.btnStop, ms=250)

    # ----- Serial batch handler -----
    # ----- PR02 ** TODO ADD METRIC -----
    def on_serial_batch(self, lines: list[str]):
        if not lines: return
        lines = lines[-400:]
        if self.chkEcho.isChecked():
            self.append_log("\n".join(lines[-40:]))

        t0 = self._t0
        for raw in lines:
            parsed = parse_p_line(raw)
            if not parsed: continue
            t_ms = int((time.monotonic() - t0) * 1000.0)
            self._latest_snapshot = (t_ms, parsed)

            # CSV (unchanged)
            if self.chkCSV.isChecked():
                self.sample += 1
                row = [self.sample, t_ms]
                for key in ("p1X","p1Y","p1R","p1S",
                            "p2X","p2Y","p2R","p2S",
                            "p3X","p3Y","p3R","p3S",
                            "p4X","p4Y","p4R","p4S"):
                    row.append(parsed.get(key, ""))
                try: self.csv.q.put_nowait(row)
                except queue.Full:
                    try: _ = self.csv.q.get_nowait()
                    except queue.Empty: pass
                    self.csv.q.put_nowait(row)

            # Plot (use selected metrics × pouches)
            if self.chkLivePlot.isChecked():
                if (t_ms - self._last_plot_emit_ms) >= self._plot_min_interval_ms:
                    self._last_plot_emit_ms = t_ms

                    # selected pouches
                    pouches = []
                    if self.chkP1.isChecked(): pouches.append(1)
                    if self.chkP2.isChecked(): pouches.append(2)
                    if self.chkP3.isChecked(): pouches.append(3)
                    if self.chkP4.isChecked(): pouches.append(4)

                    # selected metrics
                    metrics = []
                    if self.chkMX.isChecked(): metrics.append("X")
                    if self.chkMY.isChecked(): metrics.append("Y")
                    if self.chkMR.isChecked(): metrics.append("R")
                    if self.chkMS.isChecked(): metrics.append("S")

                    # protect FPS if user selects too many
                    max_series, pushed = 8, 0
                    for p in pouches:
                        for m in metrics:
                            if pushed >= max_series: break
                            key = f"p{p}{m}"
                            if key in parsed:
                                self.plot.push(f"P{p}{m}", t_ms, float(parsed[key]))
                                pushed += 1
                        if pushed >= max_series: break

    # UI timer snapshot
    def _drain_snapshot(self):
        snap = self._latest_snapshot
        if not snap: return
        t_ms, parsed = snap
        ys = [parsed.get(k) for k in ("p1Y","p2Y","p3Y","p4Y") if k in parsed]
        if ys:
            self.statusBar().showMessage(f"t={t_ms} ms  Y={ys}")

    def closeEvent(self, e):
        try:
            self.ser.disconnect()
            if hasattr(self, "csv") and self.csv: self.csv.close()
        finally:
            super().closeEvent(e)

# -------- main --------
def main():
    app = QApplication(sys.argv)
    win = MainWindow(); win.show()
    ret = app.exec_()
    try:
        if hasattr(win, "csv") and win.csv: win.csv.close()
    except Exception:
        pass
    sys.exit(ret)

if __name__ == "__main__":
    main()
