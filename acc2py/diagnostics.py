"""Coleta de avisos da transpilação (vira o RELATORIO.md)."""
from __future__ import annotations

from dataclasses import dataclass, field

LEVELS = {"info": "ℹ️", "aviso": "⚠️", "erro": "❌"}


@dataclass
class Diag:
    level: str
    where: str
    message: str


@dataclass
class Diagnostics:
    items: list[Diag] = field(default_factory=list)

    def add(self, level, where, message):
        d = Diag(level, where, message)
        if d not in self.items:
            self.items.append(d)

    def info(self, where, message):
        self.add("info", where, message)

    def warn(self, where, message):
        self.add("aviso", where, message)

    def error(self, where, message):
        self.add("erro", where, message)

    def count(self, level):
        return sum(1 for d in self.items if d.level == level)

    def to_markdown(self) -> str:
        if not self.items:
            return "Nenhum aviso.\n"
        lines = ["| Nível | Onde | Mensagem |", "|---|---|---|"]
        order = {"erro": 0, "aviso": 1, "info": 2}
        for d in sorted(self.items, key=lambda d: (order[d.level], d.where)):
            msg = d.message.replace("|", "\\|").replace("\n", " ")
            lines.append(f"| {LEVELS[d.level]} {d.level} | {d.where} | {msg} |")
        return "\n".join(lines) + "\n"
