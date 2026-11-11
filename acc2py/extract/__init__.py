"""Extração de consultas, módulos VBA e tabelas de um banco Access.

Três caminhos (``backend``):

* ``com``    — Windows com Microsoft Access instalado + ``pywin32``. É o mais
               fiel: lê o SQL de cada QueryDef e o código de cada módulo pelo
               próprio Access (macros e VBA do banco ficam desativados).
* ``parser`` — qualquer sistema, com ``pip install access-parser``. Lê as
               tabelas, reconstrói o SQL das consultas a partir de MSysQueries
               e tenta extrair o VBA de MSysAccessStorage (experimental).
* ``pasta``  — uma pasta com arquivos exportados: ``*.sql`` (uma consulta por
               arquivo), saídas do ``SaveAsText`` (consultas, módulos e
               formulários), ``*.bas``/``*.cls`` e tabelas em ``*.csv``/``*.xlsx``.

``auto`` escolhe ``pasta`` para diretórios, ``com`` no Windows quando possível
e ``parser`` nos demais casos.
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass


@dataclass
class ModuleSource:
    name: str
    kind: str  # standard | class | form | report
    code: str


class ExtractorError(RuntimeError):
    pass


def open_extractor(path, backend: str = "auto"):
    path = os.fspath(path)
    if backend == "auto":
        if os.path.isdir(path):
            backend = "pasta"
        elif sys.platform == "win32":
            try:
                import win32com.client  # noqa: F401
                backend = "com"
            except ImportError:
                backend = "parser"
        else:
            backend = "parser"
    if backend == "pasta":
        from .folder import FolderExtractor
        return FolderExtractor(path)
    if backend == "com":
        from .com import ComExtractor
        return ComExtractor(path)
    if backend == "parser":
        from .parser_backend import ParserExtractor
        return ParserExtractor(path)
    raise ExtractorError(f"backend desconhecido: {backend}")
