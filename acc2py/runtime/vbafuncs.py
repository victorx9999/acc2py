"""Funções nativas do VBA/Access em Python (versões escalares).

As mesmas funções servem ao VBA transpilado (``vb.Left(...)``) e, vetorizadas,
às consultas SQL (``F.Left(...)``). ``None`` representa o Null do Access.

O comportamento que depende das configurações regionais do Windows
(Format, CStr de números, CDate de texto) segue ``LOCALE``, que por padrão
é pt-BR. Use ``set_locale("en-US")`` se o banco foi feito em inglês.
"""
from __future__ import annotations

import math
import random
import re
from datetime import date, datetime, time, timedelta
from decimal import ROUND_HALF_EVEN, Decimal

__all__: list[str] = []
CANONICAL: dict[str, str] = {}  # nome minúsculo -> nome da função Python


def vbafunc(*aliases):
    """Registra a função como nativa do VBA (e seus apelidos, ex.: Left$)."""
    def deco(fn):
        name = fn.__name__.rstrip("_")
        __all__.append(fn.__name__)
        CANONICAL[name.lower()] = fn.__name__
        for a in aliases:
            CANONICAL[a.lower()] = fn.__name__
        return fn
    return deco


# ----------------------------------------------------------------- locale
LOCALES = {
    "pt-BR": dict(decimal=",", thousands=".", dmy=True, short_date="dd/mm/yyyy",
                  medium_date="dd-mmm-yy", long_date="dddd, d \"de\" mmmm \"de\" yyyy",
                  currency="R$ ", true="Verdadeiro", false="Falso",
                  months=["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
                          "agosto", "setembro", "outubro", "novembro", "dezembro"],
                  days=["domingo", "segunda-feira", "terça-feira", "quarta-feira",
                        "quinta-feira", "sexta-feira", "sábado"]),
    "en-US": dict(decimal=".", thousands=",", dmy=False, short_date="m/d/yyyy",
                  medium_date="dd-mmm-yy", long_date="dddd, mmmm d, yyyy",
                  currency="$", true="True", false="False",
                  months=["January", "February", "March", "April", "May", "June", "July",
                          "August", "September", "October", "November", "December"],
                  days=["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                        "Saturday"]),
}
LOCALE = dict(LOCALES["pt-BR"])
OLE_EPOCH = datetime(1899, 12, 30)


def set_locale(name: str):
    LOCALE.clear()
    LOCALE.update(LOCALES[name])


# ----------------------------------------------------------------- Null
def is_null(v) -> bool:
    if v is None:
        return True
    try:
        import pandas as pd
        if v is pd.NA or v is pd.NaT:
            return True
    except ImportError:  # pragma: no cover
        pass
    return isinstance(v, float) and math.isnan(v)


def _num(v):
    """Converte para número seguindo o VBA (texto numérico, booleano = -1/0)."""
    if is_null(v):
        return None
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, (int, float, Decimal)):
        return v
    if isinstance(v, datetime):
        return (v - OLE_EPOCH) / timedelta(days=1)
    if isinstance(v, str):
        s = v.strip()
        if LOCALE["decimal"] == ",":
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
        try:
            return int(s)
        except ValueError:
            try:
                return float(s)
            except ValueError:
                raise TypeError(f"Tipo incompatível (erro 13): {v!r}") from None
    if hasattr(v, "item"):  # escalares numpy
        return v.item()
    raise TypeError(f"Tipo incompatível (erro 13): {v!r}")


def _str(v) -> str:
    return CStr(v)


# ----------------------------------------------------------------- texto
@vbafunc("Left$")
def Left(s, n):
    return None if is_null(s) else _str(s)[: int(n)]


@vbafunc("Right$")
def Right(s, n):
    if is_null(s):
        return None
    n = int(n)
    return _str(s)[-n:] if n > 0 else ""


@vbafunc("Mid$")
def Mid(s, start, length=None):
    if is_null(s):
        return None
    s = _str(s)
    start = int(start) - 1
    return s[start:] if length is None or is_null(length) else s[start: start + int(length)]


@vbafunc("LenB")
def Len(s):
    return None if is_null(s) else len(_str(s))


@vbafunc("Trim$")
def Trim(s):
    return None if is_null(s) else _str(s).strip(" ")


@vbafunc("LTrim$")
def LTrim(s):
    return None if is_null(s) else _str(s).lstrip(" ")


@vbafunc("RTrim$")
def RTrim(s):
    return None if is_null(s) else _str(s).rstrip(" ")


@vbafunc("UCase$")
def UCase(s):
    return None if is_null(s) else _str(s).upper()


@vbafunc("LCase$")
def LCase(s):
    return None if is_null(s) else _str(s).lower()


@vbafunc()
def StrConv(s, conversion, lcid=None):
    if is_null(s):
        return None
    s = _str(s)
    return {1: s.upper(), 2: s.lower(), 3: s.title()}.get(int(conversion), s)


@vbafunc()
def InStr(*args):
    """InStr([início,] texto, procura [, compare]) — base 1, 0 se não achar."""
    if len(args) >= 3 and not isinstance(args[0], str) and not is_null(args[0]):
        start, s, sub = int(args[0]), args[1], args[2]
        compare = args[3] if len(args) > 3 else 1
    else:
        start, s, sub = 1, args[0], args[1]
        compare = args[2] if len(args) > 2 else 1
    if is_null(s) or is_null(sub):
        return None
    s, sub = _str(s), _str(sub)
    if compare != 0:  # Access: Option Compare Database -> sem diferenciar maiúsculas
        s, sub = s.lower(), sub.lower()
    return s.find(sub, start - 1) + 1


@vbafunc()
def InStrRev(s, sub, start=-1, compare=1):
    if is_null(s) or is_null(sub):
        return None
    s, sub = _str(s), _str(sub)
    if compare != 0:
        s, sub = s.lower(), sub.lower()
    end = len(s) if start == -1 else int(start)
    return s.rfind(sub, 0, end) + 1


@vbafunc()
def Replace(s, find, repl, start=1, count=-1, compare=1):
    if is_null(s):
        return None
    s = _str(s)[int(start) - 1:]
    find, repl = _str(find), _str(repl)
    if not find:
        return s
    if compare == 0:
        return s.replace(find, repl, count if count >= 0 else -1)
    return re.sub(re.escape(find), lambda _: repl, s, count=max(count, 0), flags=re.IGNORECASE)


@vbafunc("Space$")
def Space(n):
    return " " * int(n)


@vbafunc("String$")
def String(n, ch):
    ch = chr(ch) if isinstance(ch, int) else _str(ch)[0]
    return ch * int(n)


@vbafunc()
def StrReverse(s):
    return None if is_null(s) else _str(s)[::-1]


@vbafunc()
def StrComp(a, b, compare=1):
    if is_null(a) or is_null(b):
        return None
    a, b = _str(a), _str(b)
    if compare != 0:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


@vbafunc("Chr$", "ChrW", "ChrW$")
def Chr(code):
    return chr(int(code))


@vbafunc("AscW")
def Asc(s):
    return ord(_str(s)[0])


@vbafunc()
def Val(s):
    if is_null(s):
        raise TypeError("Uso inválido de Null (erro 94)")
    t = _str(s).replace(" ", "").replace("\t", "")
    m = re.match(r"&[Hh]([0-9A-Fa-f]+)", t)
    if m:
        return int(m.group(1), 16)
    m = re.match(r"&[Oo]([0-7]+)", t)
    if m:
        return int(m.group(1), 8)
    m = re.match(r"[+-]?(\d+\.?\d*|\.\d+)([eEdD][+-]?\d+)?", t)
    if not m:
        return 0
    v = float(m.group().replace("d", "e").replace("D", "e"))
    return int(v) if v.is_integer() and "." not in m.group() else v


@vbafunc("Str$")
def Str(n):
    if is_null(n):
        return None
    v = _num(n)
    txt = _fmt_general(v, ".")
    return txt if v < 0 else " " + txt


def _fmt_general(v, dec):
    if isinstance(v, float):
        if v.is_integer() and abs(v) < 1e15:
            return str(int(v))
        txt = repr(round(v, 15 - max(0, int(math.log10(abs(v))) + 1)) if v else 0.0)
        txt = txt.rstrip("0").rstrip(".") if "e" not in txt else txt.upper().replace("E", "E")
        return txt.replace(".", dec)
    if isinstance(v, Decimal):
        return format(v.normalize(), "f").replace(".", dec)
    return str(v)


@vbafunc("CStr$", "CVar")
def CStr(v):
    if isinstance(v, str):
        return v
    if is_null(v):
        return None
    if isinstance(v, bool):
        return LOCALE["true"] if v else LOCALE["false"]
    if isinstance(v, datetime):
        if v.time() == time(0) and v.date() != OLE_EPOCH.date():
            return Format(v, "Short Date")
        if v.date() == OLE_EPOCH.date():
            return Format(v, "Long Time")
        return Format(v, "Short Date") + " " + Format(v, "Long Time")
    if isinstance(v, date):
        return Format(datetime(v.year, v.month, v.day), "Short Date")
    if hasattr(v, "item"):
        v = v.item()
    return _fmt_general(v, LOCALE["decimal"])


# ----------------------------------------------------------------- números
def _bankers(v, digits=0):
    q = Decimal(1).scaleb(-digits)
    return float(Decimal(str(v)).quantize(q, rounding=ROUND_HALF_EVEN))


@vbafunc("CLng", "CByte", "CLngLng", "CLngPtr")
def CInt(v):
    n = _num(v)
    if n is None:
        raise TypeError("Uso inválido de Null (erro 94)")
    return int(_bankers(n))


@vbafunc("CSng")
def CDbl(v):
    n = _num(v)
    if n is None:
        raise TypeError("Uso inválido de Null (erro 94)")
    return float(n)


@vbafunc("CDec")
def CCur(v):
    n = _num(v)
    if n is None:
        raise TypeError("Uso inválido de Null (erro 94)")
    # Currency do VBA: 4 casas decimais. Usa float para conviver com o resto do código.
    return _bankers(n, 4)


@vbafunc()
def CBool(v):
    if isinstance(v, str):
        low = v.strip().lower()
        if low in ("true", "verdadeiro", "sim", "yes"):
            return True
        if low in ("false", "falso", "não", "nao", "no"):
            return False
    n = _num(v)
    if n is None:
        raise TypeError("Uso inválido de Null (erro 94)")
    return n != 0


@vbafunc()
def Int(v):
    n = _num(v)
    return None if n is None else math.floor(n)


@vbafunc()
def Fix(v):
    n = _num(v)
    return None if n is None else math.trunc(n)


@vbafunc()
def Abs(v):
    n = _num(v)
    return None if n is None else abs(n)


@vbafunc()
def Sgn(v):
    n = _num(v)
    return None if n is None else (n > 0) - (n < 0)


@vbafunc()
def Round(v, digits=0):
    n = _num(v)
    if n is None:
        return None
    r = _bankers(n, int(digits))
    return int(r) if int(digits) == 0 else r


@vbafunc()
def Sqr(v):
    n = _num(v)
    return None if n is None else math.sqrt(n)


@vbafunc()
def Exp(v):
    n = _num(v)
    return None if n is None else math.exp(n)


@vbafunc()
def Log(v):
    n = _num(v)
    return None if n is None else math.log(n)


@vbafunc()
def Sin(v):
    return None if is_null(v) else math.sin(_num(v))


@vbafunc()
def Cos(v):
    return None if is_null(v) else math.cos(_num(v))


@vbafunc()
def Tan(v):
    return None if is_null(v) else math.tan(_num(v))


@vbafunc()
def Atn(v):
    return None if is_null(v) else math.atan(_num(v))


@vbafunc()
def Rnd(v=None):
    return random.random()


@vbafunc()
def Randomize(seed=None):
    random.seed(seed)


@vbafunc("Hex$")
def Hex(v):
    return None if is_null(v) else format(CInt(v) & 0xFFFFFFFF if CInt(v) < 0 else CInt(v), "X")


@vbafunc("Oct$")
def Oct(v):
    return None if is_null(v) else format(CInt(v), "o")


# ----------------------------------------------------------------- Null / tipos
@vbafunc()
def Nz(v, default=None):
    if is_null(v):
        return "" if default is None else default
    return v


@vbafunc()
def IsNull(v):
    return is_null(v)


@vbafunc()
def IsEmpty(v):
    return v is None or v == ""


@vbafunc()
def IsMissing(v):
    return v is None


@vbafunc()
def IsNumeric(v):
    if is_null(v) or isinstance(v, (datetime, date)):
        return False
    try:
        _num(v)
        return True
    except (TypeError, ValueError):
        return False


@vbafunc()
def IsDate(v):
    if isinstance(v, (datetime, date)):
        return True
    try:
        CDate(v)
        return isinstance(v, str)
    except (TypeError, ValueError):
        return False


@vbafunc()
def IsArray(v):
    return isinstance(v, (list, tuple))


@vbafunc()
def IsObject(v):
    return not isinstance(v, (int, float, str, bool, datetime, Decimal, type(None)))


@vbafunc()
def IIf(cond, a, b):
    return a if (not is_null(cond) and bool(cond)) else b


@vbafunc()
def Switch(*pairs):
    for i in range(0, len(pairs) - 1, 2):
        if not is_null(pairs[i]) and pairs[i]:
            return pairs[i + 1]
    return None


@vbafunc()
def Choose(index, *options):
    i = int(_num(index))
    return options[i - 1] if 1 <= i <= len(options) else None


@vbafunc()
def TypeName(v):
    if v is None:
        return "Null"
    return {bool: "Boolean", int: "Long", float: "Double", str: "String", Decimal: "Currency",
            datetime: "Date", list: "Variant()"}.get(type(v), type(v).__name__)


@vbafunc()
def VarType(v):
    if v is None:
        return 1
    return {bool: 11, int: 3, float: 5, str: 8, Decimal: 6, datetime: 7}.get(type(v), 9)


# ----------------------------------------------------------------- datas
@vbafunc("Date$")
def Date():
    return datetime.combine(date.today(), time(0))


@vbafunc()
def Now():
    return datetime.now().replace(microsecond=0)


@vbafunc("Time$")
def Time():
    n = datetime.now()
    return datetime.combine(OLE_EPOCH.date(), n.time().replace(microsecond=0))


@vbafunc()
def Timer():
    n = datetime.now()
    return n.hour * 3600 + n.minute * 60 + n.second + n.microsecond / 1e6


def _d(v) -> datetime | None:
    if is_null(v):
        return None
    if isinstance(v, datetime):
        return v
    return CDate(v)


@vbafunc()
def Year(v):
    d = _d(v)
    return None if d is None else d.year


@vbafunc()
def Month(v):
    d = _d(v)
    return None if d is None else d.month


@vbafunc()
def Day(v):
    d = _d(v)
    return None if d is None else d.day


@vbafunc()
def Hour(v):
    d = _d(v)
    return None if d is None else d.hour


@vbafunc()
def Minute(v):
    d = _d(v)
    return None if d is None else d.minute


@vbafunc()
def Second(v):
    d = _d(v)
    return None if d is None else d.second


@vbafunc()
def Weekday(v, firstdayofweek=1):
    d = _d(v)
    if d is None:
        return None
    fdw = 1 if firstdayofweek in (0, None) else int(firstdayofweek)
    return ((d.isoweekday() % 7) + 1 - fdw) % 7 + 1


@vbafunc()
def WeekdayName(n, abbreviate=False, firstdayofweek=1):
    name = LOCALE["days"][(int(n) - 1 + int(firstdayofweek or 1) - 1) % 7]
    return name[:3] if abbreviate else name


@vbafunc()
def MonthName(n, abbreviate=False):
    name = LOCALE["months"][int(n) - 1]
    return name[:3] if abbreviate else name


def _add_months(d: datetime, n: int) -> datetime:
    m = d.month - 1 + n
    y = d.year + m // 12
    m = m % 12 + 1
    last = [31, 29 if (y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)) else 28, 31, 30, 31, 30,
            31, 31, 30, 31, 30, 31][m - 1]
    return d.replace(year=y, month=m, day=min(d.day, last))


@vbafunc()
def DateAdd(interval, n, v):
    d = _d(v)
    if d is None or is_null(n):
        return None
    n = _num(n)
    i = interval.lower()
    if i == "yyyy":
        return _add_months(d, int(n) * 12)
    if i == "q":
        return _add_months(d, int(n) * 3)
    if i == "m":
        return _add_months(d, int(n))
    if i in ("d", "y", "w"):
        return d + timedelta(days=int(n))
    if i == "ww":
        return d + timedelta(weeks=int(n))
    if i == "h":
        return d + timedelta(hours=int(n))
    if i == "n":
        return d + timedelta(minutes=int(n))
    if i == "s":
        return d + timedelta(seconds=int(n))
    raise ValueError(f"Intervalo inválido em DateAdd: {interval!r}")


@vbafunc()
def DateDiff(interval, v1, v2, firstdayofweek=1, firstweekofyear=1):
    a, b = _d(v1), _d(v2)
    if a is None or b is None:
        return None
    i = interval.lower()
    if i == "yyyy":
        return b.year - a.year
    if i == "q":
        return (b.year - a.year) * 4 + (b.month - 1) // 3 - (a.month - 1) // 3
    if i == "m":
        return (b.year - a.year) * 12 + b.month - a.month
    if i in ("d", "y"):
        return (b.date() - a.date()).days
    if i == "w":
        return int((b.date() - a.date()).days / 7)
    if i == "ww":
        fdw = int(firstdayofweek or 1)
        def week_start(d):
            back = (Weekday(d, fdw) - 1)
            return d.date() - timedelta(days=back)
        return (week_start(b) - week_start(a)).days // 7
    secs = {"h": 3600, "n": 60, "s": 1}[i]
    fa = math.floor((a - OLE_EPOCH).total_seconds() / secs)
    fb = math.floor((b - OLE_EPOCH).total_seconds() / secs)
    return fb - fa


@vbafunc()
def DatePart(interval, v, firstdayofweek=1, firstweekofyear=1):
    d = _d(v)
    if d is None:
        return None
    i = interval.lower()
    if i == "yyyy":
        return d.year
    if i == "q":
        return (d.month - 1) // 3 + 1
    if i == "m":
        return d.month
    if i == "y":
        return d.timetuple().tm_yday
    if i == "d":
        return d.day
    if i == "w":
        return Weekday(d, firstdayofweek)
    if i == "ww":
        return DateDiff("ww", datetime(d.year, 1, 1), d, firstdayofweek) + 1
    return {"h": d.hour, "n": d.minute, "s": d.second}[i]


@vbafunc()
def DateSerial(y, m, d):
    base = datetime(int(y), 1, 1)
    return _add_months(base, int(m) - 1) + timedelta(days=int(d) - 1)


@vbafunc()
def TimeSerial(h, m, s):
    return OLE_EPOCH + timedelta(hours=int(h), minutes=int(m), seconds=int(s))


@vbafunc("DateValue", "TimeValue", "CVDate")
def CDate(v):
    if is_null(v):
        raise TypeError("Uso inválido de Null (erro 94)")
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
        return OLE_EPOCH + timedelta(days=float(v))
    s = str(v).strip()
    dm = "%d/%m/%Y" if LOCALE["dmy"] else "%m/%d/%Y"
    fmts = ["%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", dm, dm + " %H:%M:%S",
            dm + " %H:%M", dm.replace("%Y", "%y"), dm.replace("/", "-"), "%H:%M:%S", "%H:%M"]
    for f in fmts:
        try:
            d = datetime.strptime(s, f)
            if f.startswith("%H"):
                d = d.replace(year=1899, month=12, day=30)
            return d
        except ValueError:
            continue
    raise TypeError(f"Tipo incompatível (erro 13): não é uma data: {v!r}")


# ----------------------------------------------------------------- Format
_NAMED = {"general date": None, "short date": "short_date", "medium date": "medium_date",
          "long date": "long_date"}


@vbafunc("Format$")
def Format(v, fmt=None, firstdayofweek=1, firstweekofyear=1):
    if is_null(v):
        if fmt and ";" in fmt and len(fmt.split(";")) >= 4:
            return _literal_only(fmt.split(";")[3])
        return ""
    if not fmt:
        return CStr(v)
    low = fmt.lower()
    if isinstance(v, str) and not low.startswith((">", "<", "@")):
        try:
            v = _num(v)
        except TypeError:
            try:
                v = CDate(v)
            except TypeError:
                return v
    if isinstance(v, (datetime, date)):
        d = v if isinstance(v, datetime) else datetime(v.year, v.month, v.day)
        if low in _NAMED:
            if _NAMED[low] is None:
                return CStr(d)
            return _format_date(d, LOCALE[_NAMED[low]])
        if low == "long time":
            return d.strftime("%H:%M:%S")
        if low == "medium time":
            return d.strftime("%I:%M %p")
        if low == "short time":
            return d.strftime("%H:%M")
        return _format_date(d, fmt)
    if isinstance(v, str):
        if low == ">":
            return v.upper()
        if low == "<":
            return v.lower()
        return v
    if isinstance(v, bool) or low in ("yes/no", "true/false", "on/off"):
        b = bool(v)
        pairs = {"yes/no": ("Sim", "Não") if LOCALE["dmy"] else ("Yes", "No"),
                 "true/false": (LOCALE["true"], LOCALE["false"]),
                 "on/off": ("Ativado", "Desativado") if LOCALE["dmy"] else ("On", "Off")}
        if low in pairs:
            return pairs[low][0 if b else 1]
        v = -1 if b else 0
    n = _num(v)
    named = {"general number": None, "currency": LOCALE["currency"] + "#,##0.00",
             "fixed": "0.00", "standard": "#,##0.00", "percent": "0.00%", "scientific": "0.00E+00"}
    if low in named:
        if named[low] is None:
            return CStr(n)
        if low == "scientific":
            return f"{float(n):.2E}".replace(".", LOCALE["decimal"])
        return _format_number(n, named[low])
    return _format_number(n, fmt)


def _literal_only(sec):
    return re.sub(r'"([^"]*)"', r"\1", sec)


def _format_number(n, fmt: str) -> str:
    sections = _split_sections(fmt)
    if n < 0 and len(sections) > 1:
        sec, n, neg = sections[1], -n, False
    elif n == 0 and len(sections) > 2:
        sec, neg = sections[2], False
    else:
        sec, neg = sections[0], n < 0
        n = abs(n)
    # separa literais do padrão numérico
    tokens = re.findall(r'"[^"]*"|\\.|[0#.,%Ee+\-]+|.', sec)
    pattern = "".join(t for t in tokens if re.fullmatch(r"[0#.,%]+", t))
    if "%" in pattern:
        n = n * 100
    num_part = pattern.replace("%", "")
    if "." in num_part:
        int_pat, dec_pat = num_part.split(".", 1)
    else:
        int_pat, dec_pat = num_part, ""
    thousands = "," in int_pat
    int_pat = int_pat.replace(",", "")
    min_dec = dec_pat.count("0")
    max_dec = len(dec_pat.replace(",", ""))
    rounded = Decimal(str(n)).quantize(Decimal(1).scaleb(-max_dec), rounding="ROUND_HALF_UP")
    txt = format(rounded, "f")
    ip, _, dp = txt.partition(".")
    dp = dp.rstrip("0") if len(dp) > min_dec else dp
    while len(dp) < min_dec:
        dp += "0"
    min_int = int_pat.count("0")
    ip = ip.lstrip("0")
    ip = ip.rjust(min_int, "0")
    if thousands and ip:
        ip = f"{int(ip):,}".replace(",", "\x00")
        ip = ip.replace("\x00", LOCALE["thousands"])
    body = ip + (LOCALE["decimal"] + dp if dp else "")
    out, placed = [], False
    for t in tokens:
        if re.fullmatch(r"[0#.,%]+", t):
            if not placed:
                out.append(body)
                placed = True
            if "%" in t:
                out.append("%")
        elif t.startswith('"'):
            out.append(t[1:-1])
        elif t.startswith("\\"):
            out.append(t[1:])
        else:
            out.append(t)
    res = "".join(out)
    return "-" + res if neg and rounded != 0 else res


def _split_sections(fmt):
    out, cur, inq = [], "", False
    for ch in fmt:
        if ch == '"':
            inq = not inq
        if ch == ";" and not inq:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def _format_date(d: datetime, fmt: str) -> str:
    tokens = re.findall(r'"[^"]*"|\\.|dddd|ddd|dd|d|mmmm|mmm|mm|m|yyyy|yy|y|hh|h|nn|n|ss|s|'
                        r'am/pm|a/p|q|ww|w|.', fmt, flags=re.IGNORECASE)
    out = []
    has_hour = any(t.lower() in ("h", "hh") for t in tokens)
    ampm = any(t.lower() in ("am/pm", "a/p") for t in tokens)
    for i, t in enumerate(tokens):
        low = t.lower()
        if t.startswith('"'):
            out.append(t[1:-1])
        elif t.startswith("\\"):
            out.append(t[1:])
        elif low == "dddd":
            out.append(LOCALE["days"][d.isoweekday() % 7])
        elif low == "ddd":
            out.append(LOCALE["days"][d.isoweekday() % 7][:3])
        elif low == "dd":
            out.append(f"{d.day:02d}")
        elif low == "d":
            out.append(str(d.day))
        elif low in ("mm", "m"):
            prev = next((x.lower() for x in reversed(tokens[:i]) if x.strip(" :/-.")), "")
            if has_hour and prev in ("h", "hh"):  # mm depois de hora = minutos
                out.append(f"{d.minute:02d}" if low == "mm" else str(d.minute))
            else:
                out.append(f"{d.month:02d}" if low == "mm" else str(d.month))
        elif low == "mmmm":
            out.append(LOCALE["months"][d.month - 1])
        elif low == "mmm":
            out.append(LOCALE["months"][d.month - 1][:3])
        elif low == "yyyy":
            out.append(f"{d.year:04d}")
        elif low == "yy":
            out.append(f"{d.year % 100:02d}")
        elif low == "y":
            out.append(str(d.timetuple().tm_yday))
        elif low in ("hh", "h"):
            h = d.hour % 12 or 12 if ampm else d.hour
            out.append(f"{h:02d}" if low == "hh" else str(h))
        elif low in ("nn", "n"):
            out.append(f"{d.minute:02d}" if low == "nn" else str(d.minute))
        elif low in ("ss", "s"):
            out.append(f"{d.second:02d}" if low == "ss" else str(d.second))
        elif low == "am/pm":
            out.append("AM" if d.hour < 12 else "PM")
        elif low == "a/p":
            out.append("A" if d.hour < 12 else "P")
        elif low == "q":
            out.append(str((d.month - 1) // 3 + 1))
        elif low == "ww":
            out.append(str(DatePart("ww", d)))
        elif low == "w":
            out.append(str(Weekday(d)))
        else:
            out.append(t)
    return "".join(out)


@vbafunc()
def FormatNumber(v, digits=2, *_):
    return _format_number(_num(v), "#,##0" + ("." + "0" * int(digits) if int(digits) > 0 else ""))


@vbafunc()
def FormatCurrency(v, digits=2, *_):
    return LOCALE["currency"] + FormatNumber(v, digits)


@vbafunc()
def FormatPercent(v, digits=2, *_):
    return _format_number(_num(v), "0" + ("." + "0" * int(digits) if int(digits) > 0 else "") + "%")


@vbafunc()
def FormatDateTime(v, named=0):
    return Format(v, {0: "General Date", 1: "Long Date", 2: "Short Date", 3: "Long Time",
                      4: "Short Time"}[int(named)])


# ----------------------------------------------------------------- vetores
@vbafunc()
def Split(s, delim=" ", limit=-1, compare=0):
    if is_null(s) or s == "":
        return []
    parts = _str(s).split(delim, limit - 1 if limit > 0 else -1)
    return parts


@vbafunc()
def Join(items, delim=" "):
    return delim.join(_str(i) for i in items)


@vbafunc()
def UBound(arr, dim=1):
    lb = getattr(arr, "lbounds", None)
    if lb is not None:
        return arr.ubound(dim)
    return len(arr) - 1


@vbafunc()
def LBound(arr, dim=1):
    lb = getattr(arr, "lbounds", None)
    if lb is not None:
        return arr.lbound(dim)
    return 0


@vbafunc()
def Array(*items):
    return list(items)
