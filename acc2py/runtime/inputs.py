"""Leitura das tabelas de entrada a partir de Excel, CSV e TXT.

Depois que o Access é desligado, cada tabela que a rotina usa vem de um
arquivo. O esquema gravado na migração (``entradas.json``: colunas e tipos do
Access) é usado para:

* ler colunas de texto como texto (preserva zeros à esquerda: CPF, agência…);
* converter datas ``dd/mm/aaaa`` e números com vírgula decimal (``1.234,56``);
* casar nomes de colunas sem diferenciar maiúsculas, acentos ou espaços extras;
* avisar quando falta uma coluna, sobra coluna ou um valor não converte.
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

TIPOS = ("texto", "numero", "inteiro", "data", "logico")
_TRUE = {"sim", "s", "true", "verdadeiro", "v", "yes", "y", "1", "-1", "x"}
_FALSE = {"não", "nao", "n", "false", "falso", "f", "no", "0", ""}


@dataclass
class Checagem:
    """Resultado da leitura de um arquivo de entrada."""
    tabela: str
    arquivo: str = ""
    linhas: int = 0
    colunas: list = field(default_factory=list)
    faltando: list = field(default_factory=list)     # colunas esperadas que não vieram
    sobrando: list = field(default_factory=list)     # colunas do arquivo que o Access não tinha
    conversoes: list = field(default_factory=list)   # (coluna, tipo, nº de valores que não converteram, exemplos)
    erro: str = ""

    @property
    def ok(self) -> bool:
        return not self.erro and not self.faltando

    def resumo(self) -> str:
        if self.erro:
            return f"ERRO: {self.erro}"
        partes = [f"{self.linhas} linha(s), {len(self.colunas)} coluna(s)"]
        if self.faltando:
            partes.append("faltam colunas: " + ", ".join(self.faltando))
        if self.sobrando:
            partes.append("colunas a mais (ignoradas pelas consultas): " + ", ".join(self.sobrando))
        for col, tipo, n, ex in self.conversoes:
            partes.append(f"{n} valor(es) de '{col}' não são {tipo} (ex.: {', '.join(map(str, ex))})")
        return "; ".join(partes)

    def to_dict(self):
        return {"tabela": self.tabela, "arquivo": self.arquivo, "linhas": self.linhas,
                "colunas": self.colunas, "faltando": self.faltando, "sobrando": self.sobrando,
                "conversoes": [list(c) for c in self.conversoes], "erro": self.erro,
                "ok": self.ok, "resumo": self.resumo()}


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^0-9a-z]+", "", s)


# ======================================================================= leitura bruta
def _decode(raw: bytes, encoding: Optional[str]) -> str:
    if encoding:
        return raw.decode(encoding)
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    try:
        return raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        return raw.decode("cp1252")  # padrão do Windows em português


def _sniff_sep(sample: str) -> str:
    lines = [ln for ln in sample.splitlines()[:20] if ln.strip()]
    best, score = ";", -1
    for sep in (";", "\t", ",", "|"):
        counts = [ln.count(sep) for ln in lines]
        if counts and min(counts) > 0 and len(set(counts)) == 1:
            s = counts[0] * 10
        elif counts:
            s = min(counts)
        else:
            s = 0
        if s > score:
            best, score = sep, s
    return best


def read_raw(path, tipo: Optional[str] = None, aba=None, separador: Optional[str] = None,
             encoding: Optional[str] = None, linha_cabecalho: int = 1) -> pd.DataFrame:
    """Lê o arquivo com todas as colunas como estão (texto em CSV/TXT)."""
    path = Path(path)
    ext = (tipo or path.suffix.lstrip(".")).lower()
    header = max(int(linha_cabecalho or 1), 1) - 1
    if ext in ("xlsx", "xlsm", "xls", "excel"):
        sheet = aba if aba not in (None, "") else 0
        df = pd.read_excel(path, sheet_name=sheet, header=header, dtype=object)
        if isinstance(df, dict):  # várias abas: usa a primeira
            df = next(iter(df.values()))
        return df
    if ext in ("csv", "txt", "tsv", "texto"):
        text = _decode(path.read_bytes(), encoding)
        sep = separador or _sniff_sep(text[:20000])
        if sep == "\\t":
            sep = "\t"
        df = pd.read_csv(io.StringIO(text), sep=sep, header=header, dtype=str,
                         keep_default_na=False, quoting=csv.QUOTE_MINIMAL, engine="python")
        df = df.replace({"": None})
        return df
    if ext == "parquet":
        return pd.read_parquet(path)
    if ext in ("pkl", "pickle"):
        return pd.read_pickle(path)
    raise ValueError(f"tipo de arquivo não suportado: .{ext} (use Excel, CSV ou TXT delimitado)")


# ======================================================================= conversões
def _to_number(s: pd.Series, inteiro=False):
    if pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
        out = s
    else:
        txt = s.map(lambda v: None if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip())
        txt = txt.str.replace(r"^R\$\s*", "", regex=True).str.replace("%", "", regex=False)
        sample = txt.dropna()
        sample = sample[sample != ""]
        br = sample.str.contains(r",\d{1,}$", regex=True).mean() if len(sample) else 0
        us_thousands = sample.str.contains(r"^\-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$", regex=True).mean() if len(sample) else 0
        if br > 0 and br >= us_thousands:  # 1.234,56
            txt = txt.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
        else:  # 1,234.56
            txt = txt.str.replace(",", "", regex=False)
        txt = txt.str.replace(r"^\((.*)\)$", r"-\1", regex=True)  # (123) contábil = -123
        out = pd.to_numeric(txt.replace({"": None}), errors="coerce")
    if inteiro and out.dropna().map(float.is_integer if out.dtype == float else (lambda _: True)).all():
        out = out.astype("Int64")
    return out


def _to_date(s: pd.Series):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    def one(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return pd.NaT
        if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year"):
            return pd.Timestamp(v)
        if isinstance(v, (int, float)):  # número serial do Excel
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(float(v), unit="D")
        t = str(v).strip()
        if not t:
            return pd.NaT
        if re.match(r"^\d{4}-\d{2}-\d{2}", t):
            return pd.to_datetime(t, errors="coerce")
        if re.match(r"^\d{8}$", t):  # 20260131
            return pd.to_datetime(t, format="%Y%m%d", errors="coerce")
        return pd.to_datetime(t, dayfirst=True, errors="coerce")
    return s.map(one).astype("datetime64[ns]")


def _to_bool(s: pd.Series):
    def one(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return None
        if isinstance(v, (bool, np.bool_)):
            return bool(v)
        t = str(v).strip().lower()
        if t in _TRUE:
            return True
        if t in _FALSE:
            return False
        return "__invalido__"
    return s.map(one)


def _to_text(s: pd.Series):
    def one(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return None
        if isinstance(v, float) and v.is_integer():
            return str(int(v))  # Excel transforma "00123" em 123.0: ao menos tira o .0
        if hasattr(v, "strftime") and not isinstance(v, str):
            return v.strftime("%d/%m/%Y") if getattr(v, "hour", 0) == 0 else v.strftime("%d/%m/%Y %H:%M:%S")
        return str(v)
    return s.map(one).astype(object)


def convert(s: pd.Series, tipo: str):
    """Converte a série para o tipo; devolve (série, valores que não converteram)."""
    before = s.notna() & (s.astype(str).str.strip() != "")
    if tipo in ("numero", "inteiro"):
        out = _to_number(s, inteiro=(tipo == "inteiro"))
    elif tipo == "data":
        out = _to_date(s)
    elif tipo == "logico":
        out = _to_bool(s)
        bad = out == "__invalido__"
        out = out.where(~bad, None)
        return out.astype(object), s[bad].astype(str).tolist()
    else:
        return _to_text(s), []
    bad = before & out.isna()
    return out, s[bad].astype(str).tolist()


def infer_type(s: pd.Series) -> str:
    """Tipo para colunas sem esquema (ou para gerar o esquema a partir de um arquivo)."""
    if pd.api.types.is_bool_dtype(s):
        return "logico"
    if pd.api.types.is_integer_dtype(s):
        return "inteiro"
    if pd.api.types.is_numeric_dtype(s):
        return "numero"
    if pd.api.types.is_datetime64_any_dtype(s):
        return "data"
    vals = s.dropna().astype(str).str.strip()
    vals = vals[vals != ""].head(500)
    if vals.empty:
        return "texto"
    if vals.str.match(r"^0\d").any():  # zero à esquerda: código, não número
        return "texto"
    if vals.str.match(r"^\-?\d+$").all():
        return "inteiro"
    if vals.str.match(r"^\-?[\d.]*\d(,\d+)?$").all() or vals.str.match(r"^\-?[\d,]*\d(\.\d+)?$").all():
        return "numero"
    if vals.str.match(r"^\d{1,2}/\d{1,2}/\d{2,4}( \d{1,2}:\d{2}(:\d{2})?)?$").all() \
            or vals.str.match(r"^\d{4}-\d{2}-\d{2}").all():
        return "data"
    return "texto"


# ======================================================================= API principal
def read_input(tabela: str, cfg: dict, base_dir=None) -> tuple[pd.DataFrame, Checagem]:
    """Lê e confere um arquivo de entrada segundo a configuração da tabela.

    cfg: {"arquivo": caminho, "tipo": "excel"|"csv"|"txt", "aba": nome, "separador": ";",
          "encoding": "cp1252", "linha_cabecalho": 1, "colunas": {"Nome": "texto", ...}}
    """
    arquivo = cfg.get("arquivo") or ""
    path = Path(arquivo)
    if base_dir is not None and not path.is_absolute():
        path = Path(base_dir) / path
    chk = Checagem(tabela, str(path))
    if not arquivo:
        chk.erro = "nenhum arquivo informado"
        return pd.DataFrame(), chk
    if not path.exists():
        chk.erro = f"arquivo não encontrado: {path}"
        return pd.DataFrame(), chk
    try:
        raw = read_raw(path, cfg.get("tipo"), cfg.get("aba"), cfg.get("separador"),
                       cfg.get("encoding"), cfg.get("linha_cabecalho", 1))
    except Exception as e:  # noqa: BLE001
        chk.erro = f"não foi possível ler: {e}"
        return pd.DataFrame(), chk
    raw = raw.dropna(how="all")  # linhas totalmente vazias (comuns no fim de planilhas)
    raw.columns = [str(c).strip() for c in raw.columns]
    raw = raw.loc[:, [c for c in raw.columns if not re.match(r"^Unnamed: \d+$", c) or raw[c].notna().any()]]
    schema: dict = cfg.get("colunas") or {}
    by_norm = {_norm(c): c for c in raw.columns}
    out = {}
    used = set()
    for col, tipo in schema.items():
        src = by_norm.get(_norm(col))
        if src is None:
            chk.faltando.append(col)
            continue
        used.add(src)
        series, bad = convert(raw[src], tipo)
        if bad:
            chk.conversoes.append((col, tipo, len(bad), bad[:3]))
        out[col] = series
    for c in raw.columns:
        if c in used:
            continue
        if schema:
            chk.sobrando.append(c)
        series, _ = convert(raw[c], infer_type(raw[c]))
        out[c] = series
    df = pd.DataFrame(out, index=raw.index).reset_index(drop=True)
    chk.linhas = len(df)
    chk.colunas = list(df.columns)
    return df, chk


def schema_from_frame(df: pd.DataFrame) -> dict:
    return {str(c): infer_type(df[c]) for c in df.columns}


ACCESS_DAO_TYPES = {1: "logico", 2: "inteiro", 3: "inteiro", 4: "inteiro", 5: "numero", 6: "numero",
                    7: "numero", 8: "data", 10: "texto", 12: "texto", 15: "texto", 16: "inteiro",
                    19: "numero", 20: "numero", 23: "data"}
