import argparse
import colorsys
import datetime
import decimal
import math
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

NBSP = " "
DECIMAL_COMMA = True

RU = {
    "mon": ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август",
            "сентябрь", "октябрь", "ноябрь", "декабрь"],
    "mon_gen": ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа",
                "сентября", "октября", "ноября", "декабря"],
    "mon_abb": ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"],
    "day": ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"],
    "day_abb": ["пн", "вт", "ср", "чт", "пт", "сб", "вс"],
}
EN = {
    "mon": ["January", "February", "March", "April", "May", "June", "July", "August",
            "September", "October", "November", "December"],
    "mon_abb": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    "day": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
    "day_abb": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
}
EN["mon_gen"] = EN["mon"]
NAMES = RU


def q(tag):
    return "{%s}%s" % (NS_MAIN, tag)


# ----------------------------------------------------------------------
#  Цвета
# ----------------------------------------------------------------------

DEFAULT_THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5",
                 "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"]

_IDX = ("000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 "
        "808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 "
        "0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF "
        "CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 "
        "FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333").split()
INDEXED = dict((i, c) for i, c in enumerate(_IDX))  # 0..7 совпадают с 8..15 после сдвига ниже
INDEXED = dict((i + 8, c) for i, c in enumerate(_IDX))
for _i in range(8):
    INDEXED[_i] = _IDX[_i]


def apply_tint(hexc, tint):
    r, g, b = [int(hexc[i:i + 2], 16) / 255.0 for i in (0, 2, 4)]
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if tint < 0:
        l = l * (1 + tint)
    else:
        l = l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return "%02X%02X%02X" % (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))


def resolve_color(el, theme):
    if el is None or el.get("auto") in ("1", "true"):
        return None
    base = None
    if el.get("rgb"):
        base = el.get("rgb")[-6:].upper()
    elif el.get("theme") is not None:
        try:
            idx = int(el.get("theme"))
            base = theme[idx] if idx < len(theme) else None
        except ValueError:
            base = None
    elif el.get("indexed") is not None:
        try:
            base = INDEXED.get(int(el.get("indexed")))
        except ValueError:
            base = None
    if base is None:
        return None
    try:
        tint = float(el.get("tint", "0") or 0)
    except ValueError:
        tint = 0
    if tint:
        base = apply_tint(base, tint)
    return base


def read_theme(z):
    try:
        root = ET.fromstring(z.read("xl/theme/theme1.xml"))
        cs = root.find(".//{%s}clrScheme" % NS_A)
        order = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6", "hlink", "folHlink"]
        lst = []
        for name in order:
            el = cs.find("{%s}%s" % (NS_A, name))
            child = el[0]
            if child.tag.endswith("srgbClr"):
                lst.append((child.get("val") or "000000").upper())
            else:
                lst.append((child.get("lastClr") or "000000").upper())
        lst[0], lst[1] = lst[1], lst[0]
        lst[2], lst[3] = lst[3], lst[2]
        return lst
    except Exception:
        return DEFAULT_THEME


# ----------------------------------------------------------------------
#  Стили
# ----------------------------------------------------------------------

BUILTIN_FMT = {
    0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
    11: "0.00E+00", 14: "dd.mm.yyyy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy",
    18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss",
    22: "dd.mm.yyyy h:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
    46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@",
}


def date_kind(code):
    if re.search(r"\[(h+|m+|s+)\]", code, re.I):
        return "t"
    s = re.sub(r'"[^"]*"', "", code)
    s = re.sub(r"\[[^\]]*\]", "", s)
    s = re.sub(r"\\.", "", s).lower()
    has_d = "d" in s or "y" in s
    has_t = "h" in s or "s" in s
    if has_d and has_t:
        return "dt"
    if has_d:
        return "d"
    if has_t:
        return "t"
    if "m" in s:
        return "d"
    return None


def _flag(el, tag):
    e = el.find(q(tag))
    if e is None:
        return None
    return e.get("val", "1") not in ("0", "false")


def parse_font(f, theme, partial=False):
    d = {}
    for tag in ("b", "i", "strike"):
        v = _flag(f, tag)
        if v is not None:
            d[tag] = v
        elif not partial:
            d[tag] = False
    u = f.find(q("u"))
    if u is not None:
        val = u.get("val", "single")
        d["u"] = None if val == "none" else ("double" if val.startswith("double") else "single")
    elif not partial:
        d["u"] = None
    e = f.find(q("sz"))
    if e is not None and e.get("val"):
        d["sz"] = float(e.get("val"))
    elif not partial:
        d["sz"] = 11.0
    e = f.find(q("color"))
    if e is not None:
        d["color"] = resolve_color(e, theme)
    elif not partial:
        d["color"] = None
    e = f.find(q("name"))
    if e is None:
        e = f.find(q("rFont"))
    if e is not None and e.get("val"):
        d["name"] = e.get("val")
    elif not partial:
        d["name"] = "Calibri"
    e = f.find(q("vertAlign"))
    if e is not None:
        d["vert"] = e.get("val") if e.get("val") in ("superscript", "subscript") else None
    elif not partial:
        d["vert"] = None
    return d


class Styles(object):
    def __init__(self):
        self.fonts = [parse_font(ET.Element("font"), DEFAULT_THEME)]
        self.fills = [None]
        self.borders = [{}]
        self.xfs = [self.default_xf()]

    @staticmethod
    def default_xf():
        return {"code": "General", "kind": None, "font": 0, "fill": 0, "border": 0,
                "h": None, "v": None, "wrap": False, "indent": 0}

    def xf(self, i):
        if 0 <= i < len(self.xfs):
            return self.xfs[i]
        return self.xfs[0]

    def visible(self, i):
        xf = self.xf(i)
        fill = self.fills[xf["fill"]] if xf["fill"] < len(self.fills) else None
        brd = self.borders[xf["border"]] if xf["border"] < len(self.borders) else {}
        return fill is not None or bool(brd)


def read_styles(z, theme):
    st = Styles()
    try:
        root = ET.fromstring(z.read("xl/styles.xml"))
    except KeyError:
        return st

    custom = {}
    nf = root.find(q("numFmts"))
    if nf is not None:
        for f in nf.findall(q("numFmt")):
            try:
                custom[int(f.get("numFmtId"))] = f.get("formatCode", "")
            except (TypeError, ValueError):
                pass

    fe = root.find(q("fonts"))
    if fe is not None:
        st.fonts = [parse_font(f, theme) for f in fe.findall(q("font"))] or st.fonts

    fl = root.find(q("fills"))
    if fl is not None:
        st.fills = []
        for fill in fl.findall(q("fill")):
            pf = fill.find(q("patternFill"))
            color = None
            if pf is not None and pf.get("patternType") == "solid":
                color = resolve_color(pf.find(q("fgColor")), theme)
            st.fills.append(color)

    be = root.find(q("borders"))
    if be is not None:
        st.borders = []
        for b in be.findall(q("border")):
            d = {}
            for side, key in (("left", "l"), ("right", "r"), ("top", "t"), ("bottom", "b")):
                e = b.find(q(side))
                if e is not None and e.get("style"):
                    d[key] = (e.get("style"), resolve_color(e.find(q("color")), theme))
            st.borders.append(d)

    cx = root.find(q("cellXfs"))
    if cx is not None:
        st.xfs = []
        for xf in cx.findall(q("xf")):
            try:
                fid = int(xf.get("numFmtId", "0"))
            except ValueError:
                fid = 0
            code = custom.get(fid) or BUILTIN_FMT.get(fid, "General")
            al = xf.find(q("alignment"))
            d = Styles.default_xf()
            d["code"] = code
            d["kind"] = date_kind(code) if code not in ("General", "@") else None
            d["font"] = int(xf.get("fontId", "0") or 0)
            d["fill"] = int(xf.get("fillId", "0") or 0)
            d["border"] = int(xf.get("borderId", "0") or 0)
            if d["font"] >= len(st.fonts):
                d["font"] = 0
            if al is not None:
                d["h"] = al.get("horizontal")
                d["v"] = al.get("vertical")
                d["wrap"] = al.get("wrapText") in ("1", "true")
                try:
                    d["indent"] = int(al.get("indent", "0") or 0)
                except ValueError:
                    d["indent"] = 0
            st.xfs.append(d)
        if not st.xfs:
            st.xfs = [Styles.default_xf()]
    return st


# ----------------------------------------------------------------------
#  Форматирование значений (как показывает Excel)
# ----------------------------------------------------------------------

def general(vtext):
    try:
        f = float(vtext)
    except (TypeError, ValueError):
        return vtext
    if math.isnan(f) or math.isinf(f):
        return vtext
    if f == int(f) and abs(f) < 1e15:
        out = str(int(f))
    else:
        out = ("%.15g" % f).replace("e", "E")
    if DECIMAL_COMMA:
        out = out.replace(".", ",")
    return out


def split_sections(code):
    secs, cur, inq, i = [], [], False, 0
    while i < len(code):
        ch = code[i]
        if ch == '"':
            inq = not inq
            cur.append(ch)
        elif ch == "\\" and not inq:
            cur.append(code[i:i + 2])
            i += 1
        elif ch == ";" and not inq:
            secs.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    secs.append("".join(cur))
    return secs


def render_numeric(mag, sec):
    toks = []
    pct = 0
    sci = None
    i, n = 0, len(sec)
    while i < n:
        ch = sec[i]
        if ch == '"':
            j = sec.find('"', i + 1)
            j = n if j < 0 else j
            toks.append(["lit", sec[i + 1:j]])
            i = j + 1
        elif ch == "\\":
            toks.append(["lit", sec[i + 1:i + 2]])
            i += 2
        elif ch == "_":
            toks.append(["lit", " "])
            i += 2
        elif ch == "*":
            i += 2
        elif ch == "[":
            j = sec.find("]", i)
            j = n - 1 if j < 0 else j
            inner = sec[i + 1:j]
            if inner.startswith("$"):
                toks.append(["lit", inner[1:].split("-")[0]])
            i = j + 1
        elif ch in "0#?.,":
            j = i
            while j < n and sec[j] in "0#?.,":
                j += 1
            toks.append(["pat", sec[i:j]])
            i = j
        elif ch == "%":
            pct += 1
            toks.append(["lit", "%"])
            i += 1
        elif ch in "Ee" and i + 1 < n and sec[i + 1] in "+-":
            j = i + 2
            while j < n and sec[j] == "0":
                j += 1
            sci = (sec[i + 1], j - i - 2)
            toks.append(["sci", ""])
            i = j
        elif ch == "@":
            i += 1
        else:
            toks.append(["lit", ch])
            i += 1

    pats = [t for t in toks if t[0] == "pat"]
    if not pats:
        return "".join(t[1] for t in toks if t[0] == "lit")
    main = max(pats, key=lambda t: len(t[1].replace(",", "")))
    p = main[1]
    if "." in p:
        ip, fp = p.split(".", 1)
    else:
        ip, fp = p, ""
    m = re.search(r",+$", ip)
    scale = len(m.group(0)) if m else 0
    ip = ip.rstrip(",")
    thousands = "," in ip
    ip_digits = ip.replace(",", "")
    fp_digits = fp.replace(",", "")
    dec = len(fp_digits)
    min_int = ip_digits.count("0")

    try:
        d = decimal.Decimal(repr(float(mag))) * (decimal.Decimal(100) ** pct)
        d = d / (decimal.Decimal(1000) ** scale)
        exp_txt = ""
        if sci is not None:
            if d == 0:
                e10 = 0
            else:
                e10 = d.adjusted() - (len(ip_digits) - 1 if len(ip_digits) > 0 else 0)
            d = d.scaleb(-e10)
            exp_txt = "E" + ("-" if e10 < 0 else ("+" if sci[0] == "+" else "")) + \
                      str(abs(e10)).zfill(max(sci[1], 1))
        qd = d.quantize(decimal.Decimal(1).scaleb(-dec), rounding=decimal.ROUND_HALF_UP)
        s = format(qd, "f")
    except (decimal.InvalidOperation, ValueError, OverflowError):
        return general(repr(mag))

    ints, _, frac = s.partition(".")
    k = len(fp_digits)
    while k > 0 and fp_digits[k - 1] in "#?" and frac[k - 1:k] == "0":
        k -= 1
    frac = frac[:k]
    ints = ints.lstrip("0")
    if len(ints) < min_int:
        ints = ints.zfill(min_int)
    if thousands and ints:
        ints = re.sub(r"(?<=\d)(?=(\d{3})+$)", NBSP, ints)
    num = ints
    if "." in p:
        num += ("," if DECIMAL_COMMA else ".") + frac
    num += exp_txt

    out = []
    for t in toks:
        if t is main:
            out.append(num)
        elif t[0] == "lit":
            out.append(t[1])
    return "".join(out)


def fmt_num(v, code):
    secs = split_sections(code)
    if len(secs) >= 3 and v == 0:
        sec, mag, minus = secs[2], 0.0, False
    elif v < 0 and len(secs) >= 2:
        sec, mag, minus = secs[1], -v, False
    else:
        sec, mag, minus = secs[0], abs(v), v < 0
    body = render_numeric(mag, sec)
    return ("-" + body) if (minus and body.strip()) else body


DATE_TOK = re.compile(
    r'\[(h+|m+|s+)\]|"([^"]*)"|\\(.)|_(.)|\*(.)|\[[^\]]*\]|'
    r'(yyyy|yy|y|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am/pm|a/p|\.0+)|(.)',
    re.I | re.S)


def fmt_date(v, code, base):
    if v < 0:
        return "########"
    sec = split_sections(code)[0]
    toks = []
    for m in DATE_TOK.finditer(sec):
        if m.group(1):
            toks.append(("el", m.group(1).lower()))
        elif m.group(2) is not None:
            toks.append(("lit", m.group(2)))
        elif m.group(3):
            toks.append(("lit", m.group(3)))
        elif m.group(4):
            toks.append(("lit", " "))
        elif m.group(6):
            toks.append(("f", m.group(6).lower()))
        elif m.group(7) is not None:
            toks.append(("lit", m.group(7)))
    sig = [i for i, t in enumerate(toks) if t[0] != "lit"]
    for k, i in enumerate(sig):
        kind, val = toks[i]
        if kind == "f" and val in ("m", "mm"):
            prev = toks[sig[k - 1]] if k > 0 else None
            nxt = toks[sig[k + 1]] if k + 1 < len(sig) else None

            def is_h(t):
                return t and ((t[0] == "f" and t[1] in ("h", "hh")) or (t[0] == "el" and t[1][0] == "h"))

            def is_s(t):
                return t and ((t[0] == "f" and t[1] in ("s", "ss")) or (t[0] == "el" and t[1][0] == "s"))
            if is_h(prev) or is_s(nxt):
                toks[i] = ("f", "min" + val)
    try:
        total = int(math.floor(v * 86400 + 0.5))
        days, sod = divmod(total, 86400)
        dt = base + datetime.timedelta(days=days)
        _ = dt.year
    except (OverflowError, ValueError):
        return general(repr(v))
    h, rem = divmod(sod, 3600)
    mi, s = divmod(rem, 60)
    ampm = any(t[0] == "f" and t[1] in ("am/pm", "a/p") for t in toks)
    has_d = any(t[0] == "f" and t[1] in ("d", "dd") for t in toks)
    hh = (h % 12 or 12) if ampm else h
    out = []
    for kind, val in toks:
        if kind == "lit":
            out.append(val)
        elif kind == "el":
            n = total // 3600 if val[0] == "h" else (total // 60 if val[0] == "m" else total)
            out.append("%0*d" % (len(val), n))
        elif val in ("yyyy",):
            out.append("%04d" % dt.year)
        elif val in ("yy", "y"):
            out.append("%02d" % (dt.year % 100))
        elif val == "m":
            out.append(str(dt.month))
        elif val == "mm":
            out.append("%02d" % dt.month)
        elif val == "mmm":
            out.append(NAMES["mon_abb"][dt.month - 1])
        elif val == "mmmm":
            out.append(NAMES["mon_gen" if has_d else "mon"][dt.month - 1])
        elif val == "mmmmm":
            out.append(NAMES["mon"][dt.month - 1][:1].upper())
        elif val == "d":
            out.append(str(dt.day))
        elif val == "dd":
            out.append("%02d" % dt.day)
        elif val == "ddd":
            out.append(NAMES["day_abb"][dt.weekday()])
        elif val == "dddd":
            out.append(NAMES["day"][dt.weekday()])
        elif val == "h":
            out.append(str(hh))
        elif val == "hh":
            out.append("%02d" % hh)
        elif val == "minm":
            out.append(str(mi))
        elif val == "minmm":
            out.append("%02d" % mi)
        elif val == "s":
            out.append(str(s))
        elif val == "ss":
            out.append("%02d" % s)
        elif val == "am/pm":
            out.append("AM" if h < 12 else "PM")
        elif val == "a/p":
            out.append("A" if h < 12 else "P")
        elif val.startswith("."):
            out.append(("," if DECIMAL_COMMA else ".") + "0" * (len(val) - 1))
    return "".join(out)


def format_value(vtext, xf, base):
    code = xf["code"]
    if code in ("General", "@", ""):
        return general(vtext)
    try:
        v = float(vtext)
    except ValueError:
        return vtext
    if math.isnan(v) or math.isinf(v):
        return vtext
    try:
        if xf["kind"]:
            return fmt_date(v, code, base)
        return fmt_num(v, code)
    except Exception:
        return general(vtext)


# ----------------------------------------------------------------------
#  Чтение листов
# ----------------------------------------------------------------------

def col_to_idx(ref):
    letters = re.match(r"\$?([A-Za-z]+)", ref).group(1).upper()
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def parse_ref(ref):
    m = re.match(r"\$?([A-Za-z]+)\$?(\d+)", ref)
    return int(m.group(2)) - 1, col_to_idx(m.group(1))


def parse_si(si, theme):
    """Возвращает список (текст, свойства_шрифта_или_None)."""
    rs = si.findall(q("r"))
    if not rs:
        t = si.find(q("t"))
        return [((t.text or "") if t is not None else "", None)]
    runs = []
    for r in rs:
        te = r.find(q("t"))
        text = (te.text or "") if te is not None else ""
        rp = r.find(q("rPr"))
        props = parse_font(rp, theme, partial=True) if rp is not None else None
        runs.append((text, props))
    return runs


def read_shared_strings(z, theme):
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [parse_si(si, theme) for si in root.findall(q("si"))]


def cell_content(c, shared, styles, base, theme):
    """Возвращает (runs, общее_выравнивание)."""
    t = c.get("t", "n")
    v = c.find(q("v"))
    vt = v.text if (v is not None and v.text is not None) else None
    if t == "s":
        try:
            return shared[int(vt)], "left"
        except (TypeError, ValueError, IndexError):
            return [], "left"
    if t == "inlineStr":
        is_ = c.find(q("is"))
        return (parse_si(is_, theme) if is_ is not None else []), "left"
    if t == "str":
        return [(vt or "", None)], "left"
    if t == "e":
        return [(vt or "", None)], "center"
    if t == "b":
        return [("ИСТИНА" if vt == "1" else "ЛОЖЬ", None)], "center"
    if t == "d":
        m = re.match(r"(\d{4})-(\d{2})-(\d{2})", vt or "")
        txt = "%s.%s.%s" % (m.group(3), m.group(2), m.group(1)) if m else (vt or "")
        return [(txt, None)], "right"
    if vt is None:
        return [], "right"
    try:
        xfi = int(c.get("s", "0"))
    except ValueError:
        xfi = 0
    return [(format_value(vt, styles.xf(xfi), base), None)], "right"


def read_sheet(z, path, shared, styles, base, theme):
    root = ET.fromstring(z.read(path))

    colinfo = {}
    cols = root.find(q("cols"))
    if cols is not None:
        for c in cols.findall(q("col")):
            try:
                mn, mx = int(c.get("min")), int(c.get("max"))
            except (TypeError, ValueError):
                continue
            w = c.get("width")
            hid = c.get("hidden") in ("1", "true")
            for i in range(mn - 1, min(mx, 16384)):
                colinfo[i] = (float(w) if w else None, hid)
    default_px = 64
    sf = root.find(q("sheetFormatPr"))
    if sf is not None and sf.get("defaultColWidth"):
        try:
            default_px = max(8, int(round(float(sf.get("defaultColWidth")) * 7)))
        except ValueError:
            pass

    cells, rowinfo = {}, {}
    sd = root.find(q("sheetData"))
    rown = -1
    if sd is not None:
        for row in sd.findall(q("row")):
            ra = row.get("r")
            rown = int(ra) - 1 if ra else rown + 1
            ht = row.get("ht")
            rowinfo[rown] = (float(ht) if ht else None, row.get("hidden") in ("1", "true"))
            coln = -1
            for c in row.findall(q("c")):
                ref = c.get("r")
                coln = col_to_idx(ref) if ref else coln + 1
                runs, al = cell_content(c, shared, styles, base, theme)
                try:
                    xfi = int(c.get("s", "0"))
                except ValueError:
                    xfi = 0
                has_text = any(t.strip() for t, _ in runs)
                if has_text or styles.visible(xfi):
                    cells[(rown, coln)] = {"runs": runs if has_text else [], "xf": xfi, "al": al}

    merges = []
    mc = root.find(q("mergeCells"))
    if mc is not None:
        for m in mc.findall(q("mergeCell")):
            ref = m.get("ref", "")
            if ":" in ref:
                a, b = ref.split(":")
                r1, c1 = parse_ref(a)
                r2, c2 = parse_ref(b)
                merges.append((r1, c1, r2, c2))
    return cells, merges, colinfo, rowinfo, default_px


def layout(cells, merges, colinfo, rowinfo, default_px, include_hidden):
    if not cells:
        return None
    r0 = min(r for r, _ in cells)
    r1 = max(r for r, _ in cells)
    c0 = min(c for _, c in cells)
    c1 = max(c for _, c in cells)
    vis_rows = [r for r in range(r0, r1 + 1)
                if include_hidden or not rowinfo.get(r, (None, False))[1]]
    vis_cols = [c for c in range(c0, c1 + 1)
                if include_hidden or not colinfo.get(c, (None, False))[1]]
    if not vis_rows or not vis_cols:
        return None
    rpos = dict((r, i) for i, r in enumerate(vis_rows))
    cpos = dict((c, i) for i, c in enumerate(vis_cols))
    grid = [[cells.get((r, c)) for c in vis_cols] for r in vis_rows]

    widths = []
    for c in vis_cols:
        w = colinfo.get(c, (None, False))[0]
        px = int(round(w * 7)) if w else default_px
        widths.append(max(px, 8) * 15)
    heights = []
    for r in vis_rows:
        ht = rowinfo.get(r, (None, False))[0]
        heights.append(int(ht * 20) if ht else None)

    span = {}  # (r,c) -> ('o', rowspan, colspan) | ('x',) | ('v', colspan)
    for (a, b, e, d) in merges:
        if a not in rpos or b not in cpos:
            continue
        rs = [rpos[r] for r in range(a, min(e, r1) + 1) if r in rpos]
        cs = [cpos[c] for c in range(b, min(d, c1) + 1) if c in cpos]
        if len(rs) == 1 and len(cs) == 1:
            continue
        for i, r in enumerate(rs):
            for j, c in enumerate(cs):
                if i == 0 and j == 0:
                    span[(r, c)] = ("o", len(rs), len(cs))
                elif i == 0:
                    span[(r, c)] = ("x",)
                elif j == 0:
                    span[(r, c)] = ("v", len(cs))
                else:
                    span[(r, c)] = ("x",)
    return {"grid": grid, "span": span, "widths": widths, "heights": heights,
            "nr": len(vis_rows), "nc": len(vis_cols)}


def read_workbook(path, include_hidden):
    z = zipfile.ZipFile(path)
    try:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = dict((r.get("Id"), r.get("Target")) for r in rels)
        theme = read_theme(z)
        styles = read_styles(z, theme)
        shared = read_shared_strings(z, theme)

        base = datetime.datetime(1899, 12, 30)
        pr = wb.find(q("workbookPr"))
        if pr is not None and pr.get("date1904") in ("1", "true"):
            base = datetime.datetime(1904, 1, 1)

        result = []
        sheets = wb.find(q("sheets"))
        for s in sheets.findall(q("sheet")):
            name = s.get("name", "Лист")
            if s.get("state", "visible") != "visible" and not include_hidden:
                continue
            target = relmap.get(s.get("{%s}id" % NS_REL))
            if not target:
                continue
            spath = target.lstrip("/") if target.startswith("/") else "xl/" + target
            try:
                data = read_sheet(z, spath, shared, styles, base, theme)
            except KeyError:
                continue
            lay = layout(data[0], data[1], data[2], data[3], data[4], include_hidden)
            if lay is None:
                print("   лист «%s» пустой — пропущен" % name)
                continue
            result.append((name, lay))
        return result, styles
    finally:
        z.close()


# ----------------------------------------------------------------------
#  Запись docx
# ----------------------------------------------------------------------

_BAD_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def esc(s):
    return escape(_BAD_XML.sub("", s))


def esc_attr(s):
    return escape(_BAD_XML.sub("", s), {'"': "&quot;"})


W_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

CONTENT_TYPES = XML_HEAD + (
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '</Types>')

ROOT_RELS = XML_HEAD + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>')

DOC_RELS = XML_HEAD + (
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>')

STYLES = XML_HEAD + (
    '<w:styles %s>'
    '<w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="ru-RU"/>'
    '</w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    '</w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
    '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar>'
    '<w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
    '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>'
    '</w:tblCellMar></w:tblPr></w:style>'
    '</w:styles>') % W_NS

BORDER_MAP = {
    "thin": ("single", 4), "hair": ("single", 2), "medium": ("single", 12),
    "thick": ("single", 18), "dashed": ("dashed", 4), "dotted": ("dotted", 4),
    "double": ("double", 6), "mediumDashed": ("dashed", 12), "dashDot": ("dotDash", 4),
    "mediumDashDot": ("dotDash", 12), "dashDotDot": ("dotDotDash", 4),
    "mediumDashDotDot": ("dotDotDash", 12), "slantDashDot": ("dotDash", 12),
}


def border_el(tag, b):
    val, sz = BORDER_MAP.get(b[0], ("single", 4))
    return '<w:%s w:val="%s" w:sz="%d" w:space="0" w:color="%s"/>' % (tag, val, sz, b[1] or "auto")


def rpr_xml(f):
    name = esc_attr(f["name"])
    s = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>' % (name, name, name, name)
    if f["b"]:
        s += "<w:b/><w:bCs/>"
    if f["i"]:
        s += "<w:i/><w:iCs/>"
    if f["strike"]:
        s += "<w:strike/>"
    if f["color"]:
        s += '<w:color w:val="%s"/>' % f["color"]
    sz = max(2, int(round(f["sz"] * 2)))
    s += '<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)
    if f["u"]:
        s += '<w:u w:val="%s"/>' % f["u"]
    if f["vert"]:
        s += '<w:vertAlign w:val="%s"/>' % f["vert"]
    return s + "</w:rPr>"


def run_xml(text, rpr):
    out = []
    for part in re.split(r"(\n|\t)", text.replace("\r\n", "\n").replace("\r", "\n")):
        if part == "\n":
            out.append("<w:r>%s<w:br/></w:r>" % rpr)
        elif part == "\t":
            out.append("<w:r>%s<w:tab/></w:r>" % rpr)
        elif part:
            out.append('<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, esc(part)))
    return "".join(out)


JC = {"left": "left", "center": "center", "centerContinuous": "center", "right": "right",
      "justify": "both", "distributed": "distribute", "fill": "left"}
VA = {"top": "top", "center": "center", "bottom": "bottom", "justify": "center",
      "distributed": "center"}


def cell_paragraph(cell, styles):
    xf = styles.xf(cell["xf"] if cell else 0)
    base_font = styles.fonts[xf["font"]]
    h = xf["h"]
    if h in (None, "general"):
        h = cell["al"] if cell else "left"
    jc = JC.get(h, "left")
    ind = ""
    if xf["indent"] and h in ("left", "right", "distributed"):
        side = "right" if h == "right" else "left"
        ind = '<w:ind w:%s="%d"/>' % (side, xf["indent"] * 135)
    mark = rpr_xml(base_font)
    ppr = ('<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>%s'
           '<w:jc w:val="%s"/>%s</w:pPr>' % (ind, jc, mark))
    runs = []
    if cell:
        for text, props in cell["runs"]:
            f = dict(base_font)
            if props:
                f.update(props)
            runs.append(run_xml(text, rpr_xml(f)))
    return "<w:p>%s%s</w:p>" % (ppr, "".join(runs))


def sect_pr(landscape):
    if landscape:
        return ('<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
                '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
                'w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>', 16838 - 1440)
    return ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" '
            'w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>', 11906 - 2268)


def table_xml(lay, styles, avail):
    grid, span, nr, nc = lay["grid"], lay["span"], lay["nr"], lay["nc"]
    widths = list(lay["widths"])
    total = sum(widths)
    if total > avail:
        k = float(avail) / total
        widths = [max(100, int(w * k)) for w in widths]

    def own(r, c, side):
        if r < 0 or c < 0 or r >= nr or c >= nc or grid[r][c] is None:
            return None
        xf = styles.xf(grid[r][c]["xf"])
        brd = styles.borders[xf["border"]] if xf["border"] < len(styles.borders) else {}
        return brd.get(side)

    out = ['<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>'
           % sum(widths)]
    out.append("<w:tblGrid>%s</w:tblGrid>" % "".join('<w:gridCol w:w="%d"/>' % w for w in widths))

    for r in range(nr):
        ht = lay["heights"][r]
        out.append("<w:tr>%s" % ('<w:trPr><w:trHeight w:val="%d" w:hRule="atLeast"/></w:trPr>' % ht
                                 if ht else ""))
        c = 0
        while c < nc:
            s = span.get((r, c))
            if s and s[0] == "x":
                c += 1
                continue
            cs, rs, vm = 1, 1, None
            if s and s[0] == "o":
                rs, cs = s[1], s[2]
                vm = "restart" if rs > 1 else None
            elif s and s[0] == "v":
                cs = s[1]
                vm = "continue"
            cell = grid[r][c]
            xf = styles.xf(cell["xf"]) if cell else styles.xf(0)

            tcpr = '<w:tcW w:w="%d" w:type="dxa"/>' % sum(widths[c:c + cs])
            if cs > 1:
                tcpr += '<w:gridSpan w:val="%d"/>' % cs
            if vm == "restart":
                tcpr += '<w:vMerge w:val="restart"/>'
            elif vm == "continue":
                tcpr += "<w:vMerge/>"

            br, rc = r + rs - 1, c + cs - 1
            top = None if vm == "continue" else (own(r, c, "t") or own(r - 1, c, "b"))
            left = own(r, c, "l") or own(r, c - 1, "r")
            bottom = own(br, c, "b") or own(br + 1, c, "t")
            right = own(r, rc, "r") or own(r, rc + 1, "l")
            bx = ""
            for tag, b in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
                if b:
                    bx += border_el(tag, b)
            if bx:
                tcpr += "<w:tcBorders>%s</w:tcBorders>" % bx

            fill = styles.fills[xf["fill"]] if xf["fill"] < len(styles.fills) else None
            if fill:
                tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % fill
            tcpr += '<w:vAlign w:val="%s"/>' % VA.get(xf["v"], "bottom")

            out.append("<w:tc><w:tcPr>%s</w:tcPr>%s</w:tc>" % (tcpr, cell_paragraph(cell, styles)))
            c += cs
        out.append("</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def build_document(sheets, styles):
    body = []
    last = len(sheets) - 1
    for i, (_name, lay) in enumerate(sheets):
        total = sum(lay["widths"])
        landscape = total > 11906 - 2268
        sect, avail = sect_pr(landscape)
        body.append(table_xml(lay, styles, avail))
        tiny = '<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>'
        if i < last:
            body.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
                        'w:lineRule="auto"/>%s%s</w:pPr></w:p>' % (tiny, sect))
        else:
            body.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
                        'w:lineRule="auto"/>%s</w:pPr></w:p>' % tiny)
            body.append(sect)
    return XML_HEAD + "<w:document %s><w:body>%s</w:body></w:document>" % (W_NS, "".join(body))


def write_docx(path, sheets, styles):
    doc = build_document(sheets, styles)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("word/document.xml", doc)
        z.writestr("word/styles.xml", STYLES)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)


# ----------------------------------------------------------------------
#  Командная строка
# ----------------------------------------------------------------------

def collect_files(args):
    files = []
    for a in args:
        if os.path.isdir(a):
            for n in sorted(os.listdir(a)):
                files.append(os.path.join(a, n))
        elif os.path.isfile(a):
            files.append(a)
        else:
            print("Не найдено: %s" % a)
    res = []
    for f in files:
        n = os.path.basename(f)
        if n.startswith("~$"):
            continue
        if n.lower().endswith((".xlsx", ".xlsm")):
            res.append(f)
        elif n.lower().endswith(".xls"):
            print("Пропущен %s: формат .xls не поддерживается, пересохраните как .xlsx" % f)
    return res


def convert(path, outdir, include_hidden, split):
    sheets, styles = read_workbook(path, include_hidden)
    if not sheets:
        print("   нет данных для конвертации")
        return []
    stem = os.path.splitext(os.path.basename(path))[0]
    out_dir = outdir or os.path.dirname(os.path.abspath(path))
    os.makedirs(out_dir, exist_ok=True)
    outs = []
    if split:
        for name, lay in sheets:
            safe = re.sub(r'[\\/:*?"<>|]', "_", name)
            out = os.path.join(out_dir, "%s_%s.docx" % (stem, safe))
            write_docx(out, [(name, lay)], styles)
            outs.append(out)
    else:
        out = os.path.join(out_dir, stem + ".docx")
        write_docx(out, sheets, styles)
        outs.append(out)
    return outs


def main():
    global DECIMAL_COMMA, NAMES
    ap = argparse.ArgumentParser(description="Excel (.xlsx) -> Word (.docx) как Ctrl+C / Ctrl+V")
    ap.add_argument("paths", nargs="*", help="файлы .xlsx и/или папки (по умолчанию — текущая папка)")
    ap.add_argument("-o", "--out", help="папка для результата")
    ap.add_argument("--split", action="store_true", help="каждый лист в отдельный .docx")
    ap.add_argument("--dot", action="store_true", help="десятичная точка вместо запятой")
    ap.add_argument("--en", action="store_true", help="английские названия месяцев и дней")
    ap.add_argument("--hidden", action="store_true", help="переносить скрытые листы/строки/столбцы")
    args = ap.parse_args()

    DECIMAL_COMMA = not args.dot
    NAMES = EN if args.en else RU
    interactive = not args.paths
    files = collect_files(args.paths or ["."])
    if not files:
        print("Файлы .xlsx не найдены.")
    ok = 0
    for f in files:
        print("Обработка: %s" % f)
        try:
            outs = convert(f, args.out, args.hidden, args.split)
            for o in outs:
                print("   -> %s" % o)
            if outs:
                ok += 1
        except zipfile.BadZipFile:
            print("   ОШИБКА: файл повреждён или это не .xlsx")
        except PermissionError:
            print("   ОШИБКА: нет доступа к файлу (возможно, открыт в Word/Excel)")
        except Exception as e:  # noqa
            print("   ОШИБКА: %s: %s" % (type(e).__name__, e))
    print("Готово: %d из %d" % (ok, len(files)))
    if interactive and sys.stdin and sys.stdin.isatty():
        try:
            input("Нажмите Enter для выхода...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
