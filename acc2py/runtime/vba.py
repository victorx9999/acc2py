"""Runtime do VBA transpilado: funções nativas, objetos do Access/DAO e utilidades.

Antes de executar código transpilado que use o banco, ligue o Database:

    from acc2py.runtime import vba as vb
    vb.set_database(db)

Objetos de interface (formulários, MsgBox, InputBox) não existem em Python;
aqui eles têm substitutos simples e configuráveis (veja ``set_msgbox``,
``Forms.set`` e ``LOG``).
"""
from __future__ import annotations

import math
import os
import re
import shutil
from datetime import datetime
from decimal import Decimal
from typing import Any, Callable, Optional

import pandas as pd

from . import vbafuncs as V
from .vbafuncs import *  # noqa: F401,F403 — Left, Mid, Format, DateAdd…
from .core import Database, like_regex

BUILTINS: dict[str, str] = dict(V.CANONICAL)
LOG: list[str] = []  # ações de interface ignoradas (OpenForm, Requery…)
_STATE: dict[str, Any] = {"db": None, "msgbox": None, "inputbox": None}


def _builtin(*aliases):
    def deco(fn):
        BUILTINS[fn.__name__.lower()] = fn.__name__
        for a in aliases:
            BUILTINS[a.lower()] = fn.__name__
        return fn
    return deco


def _log(msg):
    LOG.append(msg)


# ======================================================================= banco
def set_database(db: Database):
    _STATE["db"] = db


def _db() -> Database:
    if _STATE["db"] is None:
        raise RuntimeError("Nenhum banco ligado: chame vb.set_database(db) antes.")
    return _STATE["db"]


# ======================================================================= erros
class VBAError(Exception):
    def __init__(self, number=0, description="", source=""):
        super().__init__(f"Erro {number}: {description}")
        self.number = number
        self.description = description
        self.source = source


class GoToExit(Exception):
    """Usada para traduzir 'GoTo Sair' dentro de um bloco com tratamento de erro."""


class _Debug:
    @staticmethod
    def Print(*items):
        DebugPrint(*items)

    @staticmethod
    def Assert(cond):
        assert cond, "Debug.Assert falhou"


Debug = None  # definido abaixo (depois de DebugPrint)


class _Err:
    """Objeto Err do VBA."""

    def __init__(self):
        self.Clear()

    def Clear(self):
        self.Number = 0
        self.Description = ""
        self.Source = ""

    def Raise(self, Number, Source="", Description=""):
        raise VBAError(Number, Description or f"Erro definido pelo aplicativo ({Number})", Source)

    def set(self, exc: BaseException):
        if isinstance(exc, VBAError):
            self.Number, self.Description, self.Source = exc.number, exc.description, exc.source
        else:
            self.Number = {ZeroDivisionError: 11, TypeError: 13, IndexError: 9, KeyError: 3265,
                           OverflowError: 6, FileNotFoundError: 53}.get(type(exc), 1000)
            self.Description = str(exc)
            self.Source = type(exc).__name__

    def __bool__(self):
        return self.Number != 0


Err = _Err()


# ======================================================================= operadores
def _null(v):
    return V.is_null(v)


@_builtin()
def Concat(*parts):
    """Operador &: Null & Null = Null; Null & "x" = "x"."""
    if all(_null(p) for p in parts):
        return None
    return "".join("" if _null(p) else V.CStr(p) for p in parts)


def _vint(x):
    return int(V._bankers(V._num(x)))


def IntDiv(a, b):
    if _null(a) or _null(b):
        return None
    return int(_vint(a) / _vint(b))


def Mod(a, b):
    if _null(a) or _null(b):
        return None
    return int(math.fmod(_vint(a), _vint(b)))


def Xor(a, b):
    if _null(a) or _null(b):
        return None
    return bool(a) != bool(b)


def Like(s, pattern):
    if _null(s) or _null(pattern):
        return None
    return bool(like_regex(str(pattern)).match(V.CStr(s)))


def _fold(v):
    return v.casefold() if isinstance(v, str) else v


def _cmp(a, b, op):
    if _null(a) or _null(b):
        return None
    if isinstance(a, str) != isinstance(b, str):
        # VBA converte texto numérico ao comparar com número
        try:
            a, b = V._num(a), V._num(b)
        except TypeError:
            a, b = V.CStr(a), V.CStr(b)
    return op(_fold(a), _fold(b))


def Eq(a, b): return _cmp(a, b, lambda x, y: x == y)
def Ne(a, b): return _cmp(a, b, lambda x, y: x != y)
def Lt(a, b): return _cmp(a, b, lambda x, y: x < y)
def Gt(a, b): return _cmp(a, b, lambda x, y: x > y)
def Le(a, b): return _cmp(a, b, lambda x, y: x <= y)
def Ge(a, b): return _cmp(a, b, lambda x, y: x >= y)


def TypeOf(obj, type_name: str) -> bool:
    names = {c.__name__.lower() for c in type(obj).__mro__}
    return type_name.split(".")[-1].lower() in names


def ForRange(start, end, step=1):
    """For i = a To b Step s  (avaliado uma vez, como no VBA; aceita frações)."""
    i, end, step = V._num(start), V._num(end), V._num(step)
    if step == 0:
        raise VBAError(5, "Step 0 gera laço infinito")
    while (i <= end) if step > 0 else (i >= end):
        yield i
        i += step


# ======================================================================= arrays
class VBArray:
    """Array com limites do VBA (Dim a(1 To 5), Dim m(3, 2)…)."""

    def __init__(self, bounds=(), default=None, _data=None):
        self.lbounds = [lo for lo, _ in bounds]
        self.ubounds = [hi for _, hi in bounds]
        self.default = default
        self.data = _data if _data is not None else {}

    def _key(self, idx):
        idx = idx if isinstance(idx, tuple) else (idx,)
        idx = tuple(int(V._num(i)) for i in idx)
        if len(idx) != len(self.lbounds):
            raise VBAError(9, "Subscrito fora do intervalo (número de dimensões)")
        for i, lo, hi in zip(idx, self.lbounds, self.ubounds):
            if i < lo or i > hi:
                raise VBAError(9, f"Subscrito fora do intervalo: {i} (limites {lo} a {hi})")
        return idx

    def __getitem__(self, idx):
        k = self._key(idx)
        if k not in self.data:
            return self.default() if callable(self.default) else self.default
        return self.data[k]

    def __setitem__(self, idx, value):
        self.data[self._key(idx)] = value

    def __iter__(self):
        if len(self.lbounds) != 1:
            raise VBAError(13, "For Each só em arrays de uma dimensão")
        for i in range(self.lbounds[0], self.ubounds[0] + 1):
            yield self[i]

    def __len__(self):
        return 0 if not self.lbounds else self.ubounds[0] - self.lbounds[0] + 1

    def lbound(self, dim=1):
        return self.lbounds[dim - 1]

    def ubound(self, dim=1):
        if not self.ubounds:
            raise VBAError(9, "Array dinâmico ainda sem ReDim")
        return self.ubounds[dim - 1]

    def tolist(self):
        return list(self)

    def __repr__(self):
        return f"VBArray({list(zip(self.lbounds, self.ubounds))}, {self.data})"


def DimArray(*bounds, default=None):
    return VBArray([(0 if lo is None else int(lo), int(hi)) for lo, hi in bounds], default)


def ReDim(arr, *bounds, preserve=False, default=None):
    new = DimArray(*bounds, default=default if default is not None else getattr(arr, "default", None))
    if preserve and arr is not None:
        if isinstance(arr, VBArray):
            for k, v in arr.data.items():
                try:
                    new._key(k)
                    new.data[k] = v
                except VBAError:
                    pass
        else:
            for i, v in enumerate(arr):
                if i <= new.ubounds[0]:
                    new.data[(i,)] = v
    return new


def Erase(arr):
    if isinstance(arr, VBArray):
        arr.data.clear()
        return arr
    return []


# ======================================================================= coleções
class Collection:
    """VBA.Collection: base 1, chave opcional sem diferenciar maiúsculas."""

    def __init__(self):
        self._items: list = []
        self._keys: list = []

    def Add(self, Item, Key=None, Before=None, After=None):
        if Key is not None and Key.lower() in [k for k in self._keys if k]:
            raise VBAError(457, "Esta chave já está associada a um elemento desta coleção")
        pos = len(self._items)
        if Before is not None:
            pos = self._index(Before)
        elif After is not None:
            pos = self._index(After) + 1
        self._items.insert(pos, Item)
        self._keys.insert(pos, Key.lower() if Key else None)

    def _index(self, k):
        if isinstance(k, str):
            try:
                return self._keys.index(k.lower())
            except ValueError:
                raise VBAError(5, f"Chave inexistente na coleção: {k!r}") from None
        i = int(k) - 1
        if not 0 <= i < len(self._items):
            raise VBAError(9, "Subscrito fora do intervalo")
        return i

    def Item(self, k):
        return self._items[self._index(k)]

    __getitem__ = Item
    __call__ = Item

    def Remove(self, k):
        i = self._index(k)
        del self._items[i]
        del self._keys[i]

    @property
    def Count(self):
        return len(self._items)

    def __iter__(self):
        return iter(list(self._items))

    def __len__(self):
        return len(self._items)


class Dictionary:
    """Scripting.Dictionary (CompareMode texto = sem diferenciar maiúsculas)."""

    def __init__(self):
        self._d: dict = {}
        self.CompareMode = 0

    def _k(self, k):
        return k.casefold() if (self.CompareMode == 1 and isinstance(k, str)) else k

    def Add(self, Key, Item):
        if self._k(Key) in self._d:
            raise VBAError(457, "Esta chave já está associada a um elemento desta coleção")
        self._d[self._k(Key)] = (Key, Item)

    def Exists(self, Key):
        return self._k(Key) in self._d

    def Item(self, Key):
        if self._k(Key) not in self._d:  # como no VBA, ler chave inexistente cria a chave
            self._d[self._k(Key)] = (Key, None)
        return self._d[self._k(Key)][1]

    __getitem__ = Item
    __call__ = Item

    def __setitem__(self, Key, Item):
        self._d[self._k(Key)] = (Key, Item)

    def Remove(self, Key):
        del self._d[self._k(Key)]

    def RemoveAll(self):
        self._d.clear()

    def Keys(self):
        return [k for k, _ in self._d.values()]

    def Items(self):
        return [v for _, v in self._d.values()]

    @property
    def Count(self):
        return len(self._d)

    def __iter__(self):
        return iter(self.Keys())

    def __len__(self):
        return len(self._d)


@_builtin()
def CreateObject(progid: str, *_):
    low = progid.lower()
    if low == "scripting.dictionary":
        return Dictionary()
    if low == "scripting.filesystemobject":
        return FileSystemObject()
    raise VBAError(429, f"CreateObject({progid!r}) não tem equivalente em Python neste runtime")


class FileSystemObject:
    def FileExists(self, p): return os.path.isfile(p)
    def FolderExists(self, p): return os.path.isdir(p)
    def CreateFolder(self, p): os.makedirs(p, exist_ok=True)
    def DeleteFile(self, p, force=False): os.remove(p)
    def CopyFile(self, a, b, overwrite=True): shutil.copy(a, b)
    def GetFileName(self, p): return os.path.basename(p)
    def GetParentFolderName(self, p): return os.path.dirname(p)
    def BuildPath(self, a, b): return os.path.join(a, b)


# ======================================================================= arquivos
@_builtin()
def Dir(path="", attributes=0):
    import glob
    if path:
        Dir._iter = iter(sorted(os.path.basename(p) for p in glob.glob(path)))
    try:
        return next(Dir._iter)
    except (StopIteration, AttributeError):
        return ""


@_builtin()
def Kill(path):
    import glob
    for p in glob.glob(path):
        os.remove(p)


@_builtin()
def MkDir(path): os.mkdir(path)


@_builtin()
def RmDir(path): os.rmdir(path)


@_builtin()
def FileCopy(a, b): shutil.copy(a, b)


@_builtin()
def CurDir(): return os.getcwd()


@_builtin()
def Environ(name): return os.environ.get(name, "")


# ======================================================================= interface
def set_msgbox(fn: Optional[Callable]):
    """Define como responder MsgBox: fn(prompt, buttons, title) -> vbYes/vbNo/vbOK…"""
    _STATE["msgbox"] = fn


def set_inputbox(fn: Optional[Callable]):
    _STATE["inputbox"] = fn


@_builtin()
def MsgBox(prompt="", buttons=0, title=None, *_):
    if _STATE["msgbox"]:
        return _STATE["msgbox"](prompt, buttons, title)
    print(f"[MsgBox{' - ' + str(title) if title else ''}] {Concat(prompt)}")
    kind = int(buttons or 0) & 7
    return {0: 1, 1: 1, 2: 3, 3: 6, 4: 6, 5: 4}.get(kind, 1)  # OK/Sim por padrão


@_builtin()
def InputBox(prompt="", title=None, default="", *_):
    if _STATE["inputbox"]:
        return _STATE["inputbox"](prompt, title, default)
    return default


def DebugPrint(*items):
    print(*["" if _null(i) else V.CStr(i) for i in items])


Debug = _Debug()


# ======================================================================= DAO
class Field:
    def __init__(self, rs: "Recordset", name: str):
        self._rs = rs
        self.Name = name

    @property
    def Value(self):
        return self._rs._get(self.Name)

    @Value.setter
    def Value(self, v):
        self._rs._set(self.Name, v)

    def __repr__(self):
        return f"Field({self.Name}={self.Value!r})"


class _Fields:
    def __init__(self, rs):
        self._rs = rs

    def __call__(self, key):
        return Field(self._rs, self._rs._colname(key))

    Item = __call__
    __getitem__ = __call__

    @property
    def Count(self):
        return len(self._rs._df.columns) - (1 if "_rowid" in self._rs._df.columns else 0)

    def __iter__(self):
        return iter(Field(self._rs, c) for c in self._rs._df.columns if c != "_rowid")


class Recordset:
    """Recordset DAO sobre um DataFrame. Gravações voltam para a tabela de origem."""

    def __init__(self, df: pd.DataFrame, table: Optional[str] = None, db: Optional[Database] = None):
        self._df = df.reset_index(drop=True)
        self._table = table
        self._db = db
        self._pos = 0
        self._buffer: Optional[dict] = None
        self._mode = None
        self.NoMatch = False
        self.Fields = _Fields(self)

    # navegação
    @property
    def EOF(self):
        return self._pos >= len(self._df)

    @property
    def BOF(self):
        return self._pos < 0 or len(self._df) == 0

    @property
    def RecordCount(self):
        return len(self._df)

    @property
    def AbsolutePosition(self):
        return self._pos

    def MoveNext(self):
        self._pos += 1

    def MovePrevious(self):
        self._pos -= 1

    def MoveFirst(self):
        self._pos = 0

    def MoveLast(self):
        self._pos = len(self._df) - 1

    def Move(self, n):
        self._pos += int(n)

    def Requery(self):
        self._pos = 0

    def Close(self):
        pass

    def Clone(self):
        return Recordset(self._df, self._table, self._db)

    def GetRows(self, n=None):
        part = self._df.iloc[self._pos: None if n is None else self._pos + int(n)]
        self._pos += len(part)
        cols = [c for c in part.columns if c != "_rowid"]
        arr = VBArray([(0, len(cols) - 1), (0, len(part) - 1)])
        for j, (_, row) in enumerate(part.iterrows()):
            for i, c in enumerate(cols):
                arr[i, j] = row[c]
        return arr

    # campos
    def _colname(self, key):
        cols = [c for c in self._df.columns if c != "_rowid"]
        if isinstance(key, (int, float)) and not isinstance(key, bool):
            return cols[int(key)]
        for c in cols:
            if c.lower() == str(key).lower():
                return c
        if self._buffer is not None and self._mode == "add":
            return str(key)
        raise VBAError(3265, f"Item não encontrado nesta coleção: campo {key!r}")

    def _get(self, name):
        if self._buffer is not None and name in self._buffer:
            return self._buffer[name]
        if self.EOF or self.BOF:
            raise VBAError(3021, "Nenhum registro atual")
        v = self._df.iloc[self._pos][self._colname(name)]
        return None if _null(v) else (v.item() if hasattr(v, "item") and not isinstance(v, str) else v)

    def _set(self, name, v):
        if self._buffer is None:
            raise VBAError(3020, "Update ou CancelUpdate sem AddNew ou Edit")
        self._buffer[self._colname(name)] = v

    def __getitem__(self, key):
        return self._get(self._colname(key))

    def __setitem__(self, key, v):
        self._set(key, v)

    def __call__(self, key):
        return self[key]

    def item(self, key):
        return self.Fields(key)

    # edição
    def AddNew(self):
        self._buffer, self._mode = {}, "add"

    def Edit(self):
        if self.EOF:
            raise VBAError(3021, "Nenhum registro atual")
        self._buffer, self._mode = {}, "edit"

    def CancelUpdate(self):
        self._buffer = self._mode = None

    def Update(self):
        if self._buffer is None:
            raise VBAError(3020, "Update sem AddNew ou Edit")
        if self._table is None:
            raise VBAError(3027, "Recordset somente leitura (não vem de uma tabela)")
        db, t = self._db, self._table
        if self._mode == "add":
            from .core import apply_insert
            apply_insert(db, t, pd.DataFrame([self._buffer]))
            new = db.tables[t].iloc[[-1]].copy()
            new["_rowid"] = new.index
            self._df = pd.concat([self._df, new], ignore_index=True)
        else:
            rid = self._df.iloc[self._pos]["_rowid"]
            table = db.tables[t]
            for k, v in self._buffer.items():
                col = next(c for c in table.columns if c.lower() == k.lower())
                if table[col].dtype != object and isinstance(v, str):
                    table[col] = table[col].astype(object)
                table.loc[rid, col] = v
                self._df.loc[self._pos, col] = v
        self._buffer = self._mode = None

    def Delete(self):
        if self._table is None:
            raise VBAError(3027, "Recordset somente leitura")
        rid = self._df.iloc[self._pos]["_rowid"]
        table = self._db.tables[self._table]
        self._db.tables[self._table] = table.drop(index=rid)
        self._df = self._df.drop(index=self._pos).reset_index(drop=True)

    # busca
    def _find(self, criteria, rows):
        from .core import truth, Params, compile_sql
        sub = self._df.iloc[list(rows)]
        if sub.empty:
            return None
        key = "__find__" + criteria
        fn = self._db._sql_cache.get(key) if self._db else None
        if fn is None:
            fn = compile_sql(f"SELECT * FROM [rs_busca] WHERE {criteria}")
            if self._db:
                self._db._sql_cache[key] = fn
        tmp = Database({"rs_busca": sub.reset_index(drop=False).rename(columns={"index": "__pos"})})
        hit = fn(tmp, None)
        return None if hit.empty else int(hit["__pos"].iloc[0])

    def FindFirst(self, criteria):
        self._go(self._find(criteria, range(len(self._df))))

    def FindNext(self, criteria):
        self._go(self._find(criteria, range(self._pos + 1, len(self._df))))

    def FindLast(self, criteria):
        r = list(range(len(self._df)))[::-1]
        self._go(self._find(criteria, r))

    def FindPrevious(self, criteria):
        self._go(self._find(criteria, range(self._pos - 1, -1, -1)))

    def _go(self, pos):
        self.NoMatch = pos is None
        if pos is not None:
            self._pos = pos

    def __iter__(self):
        while not self.EOF:
            yield self
            self.MoveNext()


class QueryDef:
    def __init__(self, db: Database, name: str, sql: Optional[str] = None):
        self._db, self.Name = db, name
        self._sql = sql
        self._params: dict = {}
        self.Parameters = _Params(self._params)

    @property
    def SQL(self):
        return self._sql if self._sql else _QUERY_SQL.get(self.Name.lower(), "")

    @SQL.setter
    def SQL(self, value):
        self._sql = value

    def _run(self):
        if self._sql:
            return self._db.sql(self._sql, self._params)
        return self._db.run(self.Name, self._params)

    def Execute(self, options=None):
        return self._run()

    def OpenRecordset(self, type_=None, options=None):
        return Recordset(self._run(), None, self._db)


class _Params:
    def __init__(self, d):
        self._d = d

    def __call__(self, name):
        d = self._d

        class _P:
            @property
            def Value(self_inner):
                return d.get(name)

            @Value.setter
            def Value(self_inner, v):
                d[name] = v
        return _P()

    def __setitem__(self, name, v):
        self._d[name] = v

    def __getitem__(self, name):
        return self._d.get(name)


_QUERY_SQL: dict[str, str] = {}  # preenchido pelo pacote gerado (SQL original das consultas)


class DAODatabase:
    def __init__(self, db: Database):
        self._db = db
        self.RecordsAffected = 0

    def OpenRecordset(self, source, type_=None, options=None, lock=None):
        db = self._db
        src = str(source).strip()
        if src in db.tables:
            df = db.tables[src].copy()
            df["_rowid"] = df.index
            return Recordset(df, src, db)
        if src in db.queries:
            return Recordset(db.run(src), None, db)
        table = _single_table_select(src)
        if table and table in db.tables:
            # SELECT * FROM tabela [WHERE…] [ORDER BY…]: recordset editável
            tmp = db.tables[table].copy()
            tmp["__rid"] = tmp.index
            res = Database({table: tmp}, functions=dict(db.functions.items())).sql(src)
            return Recordset(res.rename(columns={"__rid": "_rowid"}), table, db)
        return Recordset(db.sql(src), None, db)

    def Execute(self, sql, options=None):
        res = self._db.sql(str(sql))
        self.RecordsAffected = res if isinstance(res, int) else len(res)
        return self.RecordsAffected

    def QueryDefs(self, name):
        return QueryDef(self._db, name)

    def CreateQueryDef(self, name="", sql=None):
        from .core import compile_sql
        qd = QueryDef(self._db, name, sql)
        if name and sql:
            self._db.queries[name] = compile_sql(sql)
            _QUERY_SQL[name.lower()] = sql
        return qd

    def TableDefs(self, name):
        return self._db.tables[name]

    def Close(self):
        pass

    @property
    def Name(self):
        return "acc2py"


def _single_table_select(sql: str) -> Optional[str]:
    """Nome da tabela se o SQL for 'SELECT * FROM tabela [WHERE] [ORDER BY]'."""
    from ..sql import ast as SA
    from ..sql.parser import parse
    try:
        st = parse(sql)
    except Exception:
        return None
    if isinstance(st, SA.Select) and isinstance(st.from_, SA.TableRef) and not st.group_by \
            and st.having is None and not st.distinct and len(st.columns) == 1 \
            and isinstance(st.columns[0][0], SA.Star):
        return st.from_.name
    return None


@_builtin("CodeDb")
def CurrentDb():
    return DAODatabase(_db())


class _DBEngine:
    def Workspaces(self, i=0):
        return self

    def Databases(self, i=0):
        return CurrentDb()

    def BeginTrans(self):
        _log("BeginTrans ignorado (sem transações no runtime)")

    def CommitTrans(self, *_):
        pass

    def Rollback(self):
        _log("Rollback ignorado: alterações já aplicadas nos DataFrames")


DBEngine = _DBEngine()


# ======================================================================= DLookup etc. no VBA
def _domain(fn):
    def call(expr, domain, criteria=None):
        from . import sqlfuncs as F
        return getattr(F, fn)(_db(), expr, domain, criteria)
    call.__name__ = fn
    BUILTINS[fn.lower()] = fn
    return call


DLookup, DSum, DCount, DAvg, DMin, DMax, DFirst, DLast = (
    _domain(n) for n in ("DLookup", "DSum", "DCount", "DAvg", "DMin", "DMax", "DFirst", "DLast"))


# ======================================================================= DoCmd
class _DoCmd:
    def RunSQL(self, sql, use_transaction=None):
        return _db().sql(str(sql))

    def OpenQuery(self, name, view=None, data_mode=None):
        res = _db().run(name)
        return res

    def SetWarnings(self, on):
        pass

    def Hourglass(self, on):
        pass

    def Echo(self, *a):
        pass

    def TransferSpreadsheet(self, TransferType=1, SpreadsheetType=10, TableName="", FileName="",
                            HasFieldNames=True, Range=None, *_):
        db = _db()
        if int(TransferType) == 1:  # acExport
            df = db.get(TableName)
            sheet = (Range or TableName)[:31]
            mode = "a" if os.path.exists(FileName) else "w"
            kw = {"if_sheet_exists": "replace"} if mode == "a" else {}
            with pd.ExcelWriter(FileName, engine="openpyxl", mode=mode, **kw) as w:
                df.to_excel(w, sheet_name=sheet, index=False, header=bool(HasFieldNames))
        elif int(TransferType) == 0:  # acImport
            df = pd.read_excel(FileName, sheet_name=(Range.split("!")[0] if Range else 0),
                               header=0 if HasFieldNames else None)
            if TableName in db.tables:
                db.tables[TableName] = pd.concat([db.tables[TableName], df], ignore_index=True)
            else:
                db.tables[TableName] = df
        else:
            raise VBAError(2498, "TransferSpreadsheet: acLink não suportado")

    def TransferText(self, TransferType=2, SpecificationName=None, TableName="", FileName="",
                     HasFieldNames=False, *_):
        db = _db()
        t = int(TransferType)
        if t in (2, 3):  # acExportDelim / acExportFixed
            db.get(TableName).to_csv(FileName, index=False, header=bool(HasFieldNames), sep=",")
        elif t in (0, 1):  # acImportDelim
            df = pd.read_csv(FileName, header=0 if HasFieldNames else None)
            db.tables[TableName] = pd.concat([db.tables[TableName], df], ignore_index=True) \
                if TableName in db.tables else df
        else:
            raise VBAError(2498, f"TransferText tipo {t} não suportado")

    def OutputTo(self, ObjectType, ObjectName="", OutputFormat="", OutputFile="", *_):
        df = _db().get(ObjectName)
        fmt = str(OutputFormat).lower()
        if "xls" in fmt or str(OutputFile).lower().endswith((".xlsx", ".xls")):
            df.to_excel(OutputFile, index=False)
        elif "csv" in fmt or "txt" in fmt or str(OutputFile).lower().endswith((".csv", ".txt")):
            df.to_csv(OutputFile, index=False)
        elif "htm" in fmt:
            df.to_html(OutputFile, index=False)
        else:
            raise VBAError(2282, f"OutputTo para {OutputFormat!r} não suportado (relatórios/PDF)")

    def __getattr__(self, name):
        def ignored(*args, **kw):
            _log(f"DoCmd.{name}{args} ignorado (ação de interface do Access)")
        return ignored


DoCmd = _DoCmd()


# ======================================================================= formulários
class Control:
    def __init__(self, name, value=None):
        self.Name = name
        self.Value = value
        self.Visible = True
        self.Enabled = True
        self.Locked = False
        self.Caption = name
        self.RowSource = ""

    def Requery(self):
        pass

    def SetFocus(self):
        pass


class FormBase:
    """Base dos módulos de formulário/relatório: ``self["txtNome"]`` lê o valor do controle."""

    def __init__(self, **values):
        object.__setattr__(self, "_controls", {})
        for k, v in values.items():
            self.item(k).Value = v

    def item(self, name) -> Control:
        key = name.lower()
        if key not in self._controls:
            self._controls[key] = Control(name)
        return self._controls[key]

    def __getitem__(self, name):
        return self.item(name).Value

    def __setitem__(self, name, value):
        self.item(name).Value = value

    def __call__(self, name):
        return self[name]

    def Requery(self):
        pass

    def Refresh(self):
        pass

    def Undo(self):
        pass

    @property
    def Dirty(self):
        return False

    @Dirty.setter
    def Dirty(self, v):
        pass


class _Forms:
    def __init__(self):
        self._forms: dict[str, FormBase] = {}

    def __call__(self, name):
        key = str(name).lower()
        if key not in self._forms:
            self._forms[key] = FormBase()
        return self._forms[key]

    __getitem__ = __call__

    def set(self, form_name, **values):
        """Simula um formulário aberto:  vb.Forms.set("Pedidos", txtCliente=5)"""
        f = self(form_name)
        for k, v in values.items():
            f[k] = v
        return f

    def register(self, form_name, instance: FormBase):
        self._forms[form_name.lower()] = instance


Forms = _Forms()
Reports = _Forms()


class _Application:
    def CurrentDb(self):
        return CurrentDb()

    def __getattr__(self, name):
        def ignored(*args, **kw):
            _log(f"Application.{name}{args} ignorado")
        return ignored


Application = _Application()


# ======================================================================= constantes
CONSTANTS: dict[str, Any] = {
    "vbCrLf": "\r\n", "vbCr": "\r", "vbLf": "\n", "vbNewLine": "\r\n", "vbTab": "\t",
    "vbNullString": "", "vbNullChar": "\0", "vbBack": "\b",
    "vbOKOnly": 0, "vbOKCancel": 1, "vbAbortRetryIgnore": 2, "vbYesNoCancel": 3, "vbYesNo": 4,
    "vbRetryCancel": 5, "vbCritical": 16, "vbQuestion": 32, "vbExclamation": 48,
    "vbInformation": 64, "vbDefaultButton1": 0, "vbDefaultButton2": 256, "vbDefaultButton3": 512,
    "vbOK": 1, "vbCancel": 2, "vbAbort": 3, "vbRetry": 4, "vbIgnore": 5, "vbYes": 6, "vbNo": 7,
    "vbBinaryCompare": 0, "vbTextCompare": 1, "vbDatabaseCompare": 2,
    "vbSunday": 1, "vbMonday": 2, "vbTuesday": 3, "vbWednesday": 4, "vbThursday": 5,
    "vbFriday": 6, "vbSaturday": 7, "vbUseSystemDayOfWeek": 0, "vbFirstJan1": 1,
    "vbGeneralDate": 0, "vbLongDate": 1, "vbShortDate": 2, "vbLongTime": 3, "vbShortTime": 4,
    "vbUpperCase": 1, "vbLowerCase": 2, "vbProperCase": 3,
    "vbEmpty": 0, "vbNull": 1, "vbInteger": 2, "vbLong": 3, "vbSingle": 4, "vbDouble": 5,
    "vbCurrency": 6, "vbDate": 7, "vbString": 8, "vbObject": 9, "vbBoolean": 11, "vbArray": 8192,
    "vbObjectError": -2147221504, "vbTrue": -1, "vbFalse": 0,
    "dbOpenTable": 1, "dbOpenDynaset": 2, "dbOpenSnapshot": 4, "dbOpenForwardOnly": 8,
    "dbOpenDynamic": 16, "dbAppendOnly": 8, "dbSeeChanges": 512, "dbFailOnError": 128,
    "dbReadOnly": 4, "dbDenyWrite": 1, "dbOptimistic": 3, "dbPessimistic": 2,
    "dbText": 10, "dbLong": 4, "dbInteger": 3, "dbDouble": 7, "dbDate": 8, "dbBoolean": 1,
    "dbCurrency": 5, "dbMemo": 12,
    "acTable": 0, "acQuery": 1, "acForm": 2, "acReport": 3, "acMacro": 4, "acModule": 5,
    "acNormal": 0, "acViewNormal": 0, "acViewDesign": 1, "acViewPreview": 2, "acViewReport": 5,
    "acFormAdd": 0, "acFormEdit": 1, "acFormReadOnly": 2, "acFormPropertySettings": -1,
    "acWindowNormal": 0, "acHidden": 1, "acIcon": 2, "acDialog": 3,
    "acSaveNo": 2, "acSavePrompt": 0, "acSaveYes": 1,
    "acImport": 0, "acExport": 1, "acLink": 2,
    "acImportDelim": 0, "acImportFixed": 1, "acExportDelim": 2, "acExportFixed": 3,
    "acExportMerge": 4, "acLinkDelim": 5, "acLinkFixed": 6,
    "acSpreadsheetTypeExcel3": 0, "acSpreadsheetTypeExcel8": 8, "acSpreadsheetTypeExcel9": 8,
    "acSpreadsheetTypeExcel12": 9, "acSpreadsheetTypeExcel12Xml": 10,
    "acOutputTable": 0, "acOutputQuery": 1, "acOutputForm": 2, "acOutputReport": 3,
    "acFormatXLSX": "Excel Workbook (*.xlsx)", "acFormatXLS": "Microsoft Excel (*.xls)",
    "acFormatTXT": "MS-DOS Text (*.txt)", "acFormatHTML": "HTML (*.html)",
    "acFormatPDF": "PDF Format (*.pdf)", "acFormatRTF": "Rich Text Format (*.rtf)",
    "acCmdSaveRecord": 97, "acCmdUndo": 292, "acCmdDeleteRecord": 223,
    "acNewRec": 5, "acFirst": 2, "acLast": 3, "acNext": 1, "acPrevious": 0, "acGoTo": 4,
    "acDataErrContinue": 0, "acDataErrDisplay": 1,
}
globals().update(CONSTANTS)

OBJECTS = {"docmd": "DoCmd", "forms": "Forms",
           "reports": "Reports", "err": "Err", "application": "Application", "dbengine": "DBEngine",
           "screen": "Application"}

for _n in ("Concat", "IntDiv", "Mod", "Xor", "Like", "DebugPrint"):
    BUILTINS.setdefault(_n.lower(), _n)


def __getattr__(name):  # constantes não mapeadas: erro claro
    raise AttributeError(f"acc2py.runtime.vba não define {name!r}. Se for uma constante do "
                         f"Access/VBA, acrescente-a em CONSTANTS.")
