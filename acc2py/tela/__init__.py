"""Tela local do acc2py (abre no navegador; tudo roda nesta máquina).

* Migração (``acc2py --tela``): abrir o banco, escolher as consultas e a ordem,
  definir as tabelas de entrada, ver o código gerado e transpilar.
* Execução (``python -m <pacote>``): escolher os arquivos de entrada, preencher
  os parâmetros, executar a sequência e baixar a saída.

Usa só a biblioteca padrão do Python. O servidor escuta apenas em 127.0.0.1 e
exige um token aleatório em cada chamada, para que outras páginas abertas no
navegador não consigam acioná-lo.
"""
from __future__ import annotations

import importlib
import json
import re
import secrets
import sys
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qs, unquote, urlparse

HTML = Path(__file__).with_name("index.html")


class Estado:
    def __init__(self, origem=None, saida=None, pacote=None):
        self.origem = origem
        self.saida = saida or "saida_acc2py"
        self.pacote = pacote          # runtime.pacote.Pacote
        self.ex = None                # extrator aberto (migração)
        self.ex_origem = None
        self.consultas = {}           # nome -> sql
        self.info = {}                # nome -> {tipo, deps, tabelas, params}
        self.job = None
        self.lock = threading.Lock()
        self.modo_inicial = "execucao" if pacote is not None else "migracao"


# ======================================================================= migração
def _abrir_banco(st: Estado, origem: str, backend="auto"):
    from ..extract import open_extractor
    from ..sql.parser import parse
    from ..sql.to_pandas import SQLCompiler
    from ..sql import ast as SA
    if st.ex is not None:
        try:
            st.ex.close()
        except Exception:  # noqa: BLE001
            pass
    st.ex = open_extractor(origem, backend)
    st.ex_origem = origem
    st.consultas = st.ex.queries()
    try:
        tabelas = st.ex.table_names()
    except Exception:  # noqa: BLE001
        tabelas = []
    try:
        vinculadas = st.ex.linked_tables()
    except Exception:  # noqa: BLE001
        vinculadas = {}
    lower = {q.lower(): q for q in st.consultas}
    st.info = {}
    for nome, sql in st.consultas.items():
        item = {"nome": nome, "tipo": "?", "deps": [], "tabelas": [], "params": [], "cria": None, "erro": ""}
        try:
            stmt = parse(sql)
            if isinstance(stmt, SA.Select) and stmt.into:
                item["tipo"], item["cria"] = "criar tabela", stmt.into
            elif isinstance(stmt, (SA.Select, SA.Union, SA.Transform)):
                item["tipo"] = "seleção"
            else:
                item["tipo"] = {SA.Update: "atualização", SA.Delete: "exclusão",
                                SA.Insert: "acréscimo"}.get(type(stmt), "ação")
            item["params"] = [p for p, _ in stmt.parameters]
            comp = SQLCompiler()
            comp.compile_function(nome, sql)
            item["deps"] = sorted(lower[d.lower()] for d in comp.deps if d.lower() in lower)
            item["tabelas"] = sorted(d for d in comp.deps if d.lower() not in lower)
        except Exception as e:  # noqa: BLE001
            item["erro"] = str(e)
        st.info[nome] = item
    try:
        modulos = [{"nome": m.name, "tipo": m.kind} for m in st.ex.modules()]
    except Exception as e:  # noqa: BLE001
        modulos = []
    return {"origem": origem, "consultas": list(st.info.values()), "tabelas": tabelas,
            "vinculadas": vinculadas, "modulos": modulos}


def _entradas_previstas(st: Estado, escolhidas: list[str]):
    """Tabelas que a seleção lê e que não são criadas por ela."""
    fila, vistas = list(escolhidas), []
    while fila:
        n = fila.pop(0)
        if n in vistas or n not in st.info:
            continue
        vistas.append(n)
        fila += st.info[n]["deps"]
    criadas = {st.info[n]["cria"].lower() for n in vistas if st.info[n]["cria"]}
    try:
        reais = {t.lower(): t for t in st.ex.table_names()}
        vinc = {k.lower(): v for k, v in st.ex.linked_tables().items()}
    except Exception:  # noqa: BLE001
        reais, vinc = {}, {}
    out = {}
    for n in vistas:
        for t in st.info[n]["tabelas"]:
            k = t.lower()
            if k in criadas or (reais and k not in reais):
                continue
            nome = reais.get(k, t)
            e = out.setdefault(nome, {"tabela": nome, "vinculada": vinc.get(k, ""), "usada_por": [],
                                      "modo": "mensal" if k in vinc else "fixa"})
            if n not in e["usada_por"]:
                e["usada_por"].append(n)
    automaticas = [n for n in vistas if n not in escolhidas]
    return {"entradas": list(out.values()), "automaticas": automaticas}


def _previa(st: Estado, nome: str):
    from ..diagnostics import Diagnostics
    from ..sql.to_pandas import SQLCompiler
    sql = st.consultas[nome]
    d = Diagnostics()
    try:
        code = SQLCompiler(d).compile_function(nome, sql)
    except Exception as e:  # noqa: BLE001
        code = f"# Não convertida: {e}"
        d.error(nome, str(e))
    return {"nome": nome, "sql": sql, "python": code,
            "avisos": [{"nivel": i.level, "mensagem": i.message} for i in d.items]}


def _transpilar(st: Estado, req: dict):
    from ..project import transpile_database
    origem = req.get("origem") or st.ex_origem
    if st.ex is not None:  # libera o arquivo (o Access não abre duas vezes por COM)
        try:
            st.ex.close()
        except Exception:  # noqa: BLE001
            pass
        st.ex = None
    seq = req.get("sequencia") or []
    soltas = req.get("consultas") or []
    saida = req.get("saida") or st.saida
    res = transpile_database(origem, saida, queries=soltas or None, sequence=seq or None,
                             modules=req.get("modulos") or None, package=req.get("pacote") or None,
                             input_modes=req.get("modos") or {}, modern=bool(req.get("moderno")))
    seq_file = Path(saida) / f"sequencia_{res.package_dir.name}.txt"
    seq_file.write_text("# Ordem de execução (use com: acc2py ... -s este_arquivo)\n" + "\n".join(seq) + "\n",
                        encoding="utf-8")
    d = res.diags
    return {"pacote": str(res.package_dir), "consultas": {k: v[1] for k, v in res.queries.items()},
            "modulos": {k: v[1] for k, v in res.modules.items()}, "erros": d.count("erro"),
            "moderno": res.modern,
            "avisos": d.count("aviso"),
            "itens": [{"nivel": i.level, "onde": i.where, "mensagem": i.message} for i in d.items],
            "relatorio": str(res.package_dir / "RELATORIO.md")}


def _usar_pacote(st: Estado, pasta: str):
    p = Path(pasta).resolve()
    if not (p / "__init__.py").exists() or not (p / "consultas.py").exists():
        raise ValueError(f"não é um pacote gerado pelo acc2py: {p}")
    if str(p.parent) not in sys.path:
        sys.path.insert(0, str(p.parent))
    for k in [k for k in sys.modules if k == p.name or k.startswith(p.name + ".")]:
        del sys.modules[k]
    mod = importlib.import_module(p.name)
    st.pacote = mod.pacote()
    return {"ok": True, "pacote": st.pacote.nome}


# ======================================================================= execução
def _info_pacote(st: Estado):
    from ..runtime.pacote import from_json_value
    pk = st.pacote
    exe = pk.execucao()
    ultimos = {k.lower(): from_json_value(v) for k, v in exe.get("parametros", {}).items()}
    params = []
    for p in pk.parametros_declarados():
        v = ultimos.get(p["nome"].lower())
        if hasattr(v, "strftime"):
            v = v.strftime("%Y-%m-%d")
        params.append({**p, "valor": v})
    for k, v in exe.get("parametros", {}).items():  # extras (não declarados) usados antes
        if k.lower() not in {p["nome"].lower() for p in params}:
            v = from_json_value(v)
            params.append({"nome": k, "tipo": "auto", "tipo_access": "não declarado",
                           "valor": v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else v})
    return {"nome": pk.nome, "pasta": str(pk.dir),
            "sequencia": [{"nome": n, "tipo": pk.tipos.get(n, "?")} for n in pk.sequencia],
            "entradas": pk.situacao_entradas(), "parametros": params,
            "ultima_execucao": exe.get("ultima_execucao")}


def _executar(st: Estado, req: dict):
    from ..runtime.pacote import EntradaError, convert_param
    pk = st.pacote
    tipos = {p["nome"].lower(): p["tipo"] for p in pk.parametros_declarados()}
    params = {}
    for k, v in (req.get("parametros") or {}).items():
        if v in (None, ""):
            continue
        params[k] = convert_param(v, tipos.get(k.lower(), "auto"))
    job = {"rodando": True, "log": [], "resultado": None}
    st.job = job

    def run():
        try:
            job["log"].append("Lendo as tabelas de entrada…")
            try:
                db = pk.abrir()
            except EntradaError as e:
                job["resultado"] = {"ok": False, "erro": str(e),
                                    "checagens": [c.to_dict() for c in e.checagens]}
                return
            for c in db.checagens:
                job["log"].append(f"  {c.tabela}: {c.resumo()}")
            job["log"].append("Executando a sequência:")
            etapas, destino = pk.executar(params, intermediarias=bool(req.get("intermediarias")),
                                          log=job["log"].append, db=db)
            rows = etapas.fillna("").to_dict("records")
            falta = None
            for r in rows:
                m = re.search(r"Parâmetro não informado: '(.+?)'", str(r.get("Erro", "")))
                if m:
                    falta = m.group(1)
            ok = all(r["Situação"] == "ok" for r in rows)
            previa = None
            if ok:
                import pandas as pd
                sheets = pd.read_excel(destino, sheet_name=None, nrows=200)
                first = next((k for k in sheets if k != "Etapas"), None)
                if first:
                    df = sheets[first]
                    previa = {"aba": first, "colunas": [str(c) for c in df.columns],
                              "linhas": df.astype(object).where(df.notna(), "").astype(str).values.tolist()}
            job["resultado"] = {"ok": ok, "etapas": rows, "arquivo": Path(destino).name,
                                "caminho": str(destino), "falta_parametro": falta, "previa": previa}
        except Exception as e:  # noqa: BLE001
            job["resultado"] = {"ok": False, "erro": f"{e}", "detalhe": traceback.format_exc()}
        finally:
            job["rodando"] = False

    threading.Thread(target=run, daemon=True).start()
    return {"iniciado": True}


def _procurar(tipo: str, tipos_arquivo: Optional[list] = None):
    """Abre a janela de 'Abrir arquivo' do sistema (nesta máquina)."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        raise ValueError("janela de seleção indisponível (tkinter não instalado); digite o caminho")
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        if tipo == "pasta":
            caminho = filedialog.askdirectory()
        else:
            ft = tipos_arquivo or [("Todos", "*.*")]
            caminho = filedialog.askopenfilename(filetypes=ft)
    finally:
        root.destroy()
    return {"caminho": caminho or ""}


# ======================================================================= HTTP
def _handler(st: Estado, token: str):
    class H(BaseHTTPRequestHandler):
        def log_message(self, *a):  # silencioso
            pass

        def _send(self, code, body: bytes, ctype="application/json; charset=utf-8", extra=None):
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            self.wfile.write(body)

        def _json(self, data, code=200):
            self._send(code, json.dumps(data, ensure_ascii=False, default=str).encode("utf-8"))

        def _auth(self, qs) -> bool:
            if self.headers.get("X-Token") == token or qs.get("t", [""])[0] == token:
                return True
            self._json({"erro": "token inválido"}, 403)
            return False

        def do_GET(self):
            u = urlparse(self.path)
            qs = parse_qs(u.query)
            if u.path in ("/", "/index.html"):
                html = HTML.read_text(encoding="utf-8")
                return self._send(200, html.encode("utf-8"), "text/html; charset=utf-8")
            if not self._auth(qs):
                return
            try:
                if u.path == "/api/estado":
                    return self._json({"modo": st.modo_inicial, "origem": st.origem or "",
                                       "saida": st.saida, "tem_pacote": st.pacote is not None,
                                       "pacote": st.pacote.nome if st.pacote else ""})
                if u.path == "/api/pacote":
                    return self._json(_info_pacote(st))
                if u.path == "/api/progresso":
                    return self._json(st.job or {"rodando": False, "log": [], "resultado": None})
                if u.path == "/api/baixar":
                    nome = Path(unquote(qs.get("arquivo", [""])[0])).name
                    f = st.pacote.dir / "saidas" / nome
                    if not f.exists():
                        return self._json({"erro": "arquivo não encontrado"}, 404)
                    return self._send(200, f.read_bytes(),
                                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                      {"Content-Disposition": f'attachment; filename="{nome}"'})
            except Exception as e:  # noqa: BLE001
                return self._json({"erro": str(e)}, 500)
            self._json({"erro": "não encontrado"}, 404)

        def do_POST(self):
            u = urlparse(self.path)
            qs = parse_qs(u.query)
            if not self._auth(qs):
                return
            n = int(self.headers.get("Content-Length") or 0)
            body = self.rfile.read(n) if n else b""
            try:
                if u.path == "/api/entrada_upload":
                    tabela = unquote(qs["tabela"][0])
                    nome = Path(unquote(qs["arquivo"][0])).name
                    pasta = st.pacote.dir / "entradas"
                    pasta.mkdir(exist_ok=True)
                    seguro = re.sub(r'[<>:"/\\|?*]', "_", f"{tabela}__{nome}")
                    destino = pasta / seguro
                    destino.write_bytes(body)
                    chk = st.pacote.definir_entrada(tabela, str(destino))
                    return self._json({"checagem": chk.to_dict()})
                req = json.loads(body or b"{}")
                with st.lock:
                    if u.path == "/api/abrir":
                        return self._json(_abrir_banco(st, req["origem"], req.get("backend", "auto")))
                    if u.path == "/api/entradas_previstas":
                        return self._json(_entradas_previstas(st, req.get("consultas", [])))
                    if u.path == "/api/previa":
                        return self._json(_previa(st, req["nome"]))
                    if u.path == "/api/transpilar":
                        return self._json(_transpilar(st, req))
                    if u.path == "/api/usar_pacote":
                        return self._json(_usar_pacote(st, req["pasta"]))
                    if u.path == "/api/entrada_caminho":
                        opcoes = req.get("opcoes") or {}
                        chk = st.pacote.definir_entrada(req["tabela"], req["caminho"], **opcoes)
                        return self._json({"checagem": chk.to_dict()})
                    if u.path == "/api/entrada_opcoes":
                        cfg = st.pacote.execucao()["entradas"].get(req["tabela"], {})
                        if not cfg.get("arquivo"):
                            raise ValueError("escolha o arquivo antes das opções")
                        chk = st.pacote.definir_entrada(req["tabela"], cfg["arquivo"], **(req.get("opcoes") or {}))
                        return self._json({"checagem": chk.to_dict()})
                    if u.path == "/api/executar":
                        if st.job and st.job.get("rodando"):
                            raise ValueError("já existe uma execução em andamento")
                        return self._json(_executar(st, req))
                if u.path == "/api/procurar":
                    return self._json(_procurar(req.get("tipo", "arquivo"), req.get("filtros")))
            except Exception as e:  # noqa: BLE001
                return self._json({"erro": str(e), "detalhe": traceback.format_exc()}, 400)
            self._json({"erro": "não encontrado"}, 404)
    return H


def servir(origem=None, saida=None, pacote=None, porta: int = 0, abrir_navegador: bool = True,
           bloquear: bool = True):
    """Sobe a tela em http://127.0.0.1:<porta>/ e abre o navegador."""
    st = Estado(origem, saida, pacote)
    token = secrets.token_urlsafe(16)
    srv = ThreadingHTTPServer(("127.0.0.1", porta), _handler(st, token))
    url = f"http://127.0.0.1:{srv.server_address[1]}/?t={token}"
    print(f"Tela do acc2py em: {url}")
    print("Deixe esta janela aberta enquanto usa a tela. Ctrl+C para encerrar.")
    if abrir_navegador:
        threading.Timer(0.5, lambda: webbrowser.open(url)).start()
    if not bloquear:
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        return srv, url, st
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()
    return 0
