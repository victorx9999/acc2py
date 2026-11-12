"""Linha de comando do acc2py.

Exemplos:
    acc2py vendas.accdb -o saida
    acc2py vendas.accdb -o saida --consultas "Vendas por Mês" "Clientes Ativos"
    acc2py vendas.accdb -o saida --consultas "Vendas por Mês" --modulos todos
    acc2py pasta_exportada/ -o saida --sem-vba
    acc2py vendas.accdb -o saida --consultas-arquivo lista.txt --resultado resultados.xlsx
    acc2py vendas.accdb -o saida --consultas "Rel*" --resultado res.xlsx --param "Ano=2026"
    acc2py --listar vendas.accdb
"""
from __future__ import annotations

import argparse
import os
import sys

from .project import transpile_database


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="acc2py",
        description="Transpila consultas SQL e módulos VBA do Microsoft Access para Python/pandas.")
    ap.add_argument("origem", nargs="?", help="arquivo .accdb/.mdb ou pasta com objetos exportados")
    ap.add_argument("-o", "--saida", default="saida_acc2py", help="pasta de saída (padrão: saida_acc2py)")
    ap.add_argument("-c", "--consultas", nargs="+", metavar="NOME",
                    help="só estas consultas (e as que elas usam). Sem esta opção: todas")
    ap.add_argument("-f", "--consultas-arquivo", metavar="ARQUIVO",
                    help="arquivo .txt com um nome de consulta por linha (aceita * e ?; # comenta)")
    ap.add_argument("-s", "--sequencia", metavar="ARQUIVO",
                    help="arquivo com as consultas NA ORDEM em que rodam (uma por linha). Com "
                         "--resultado, executa as etapas em ordem (inclusive as de ação) e grava "
                         "a saída final")
    ap.add_argument("--mensais", nargs="+", metavar="TABELA",
                    help="tabelas de entrada que mudam a cada execução (vêm de arquivo). As demais "
                         "locais ficam 'fixas' (cópia do Access). Vinculadas já são mensais")
    ap.add_argument("--todas-mensais", action="store_true",
                    help="todas as tabelas de entrada vêm de arquivo a cada execução")
    ap.add_argument("--validar", metavar="PASTA_DO_PACOTE",
                    help="roda a SEQUENCIA do pacote no Access (numa cópia do banco) e no Python com as "
                         "mesmas entradas e compara. Grava o relatório em --resultado (padrão validacao.xlsx)")
    ap.add_argument("--tela", action="store_true",
                    help="abre a tela interativa (migração) no navegador")
    ap.add_argument("--etapas-intermediarias", action="store_true",
                    help="com --sequencia: grava também o resultado de cada etapa de seleção")
    ap.add_argument("-r", "--resultado", metavar="DESTINO",
                    help="depois de transpilar, executa as consultas escolhidas e grava tudo em "
                         "DESTINO (.xlsx = uma aba por consulta; outro = pasta de .csv)")
    ap.add_argument("-p", "--param", action="append", default=[], metavar="NOME=VALOR",
                    help="parâmetro para --resultado (repita para vários). Ex.: \"Data inicial=2026-01-31\"")
    ap.add_argument("-m", "--modulos", nargs="+", metavar="NOME",
                    help="módulos VBA a converter ('todos' para todos). Com --consultas e sem "
                         "--modulos, só entram os módulos com funções usadas pelas consultas")
    ap.add_argument("--sem-vba", action="store_true", help="não converter VBA")
    ap.add_argument("--backend", choices=["auto", "com", "parser", "pasta"], default="auto")
    ap.add_argument("--exportar-dados", action="store_true",
                    help="copia as tabelas para <pacote>/dados (parquet ou pickle)")
    ap.add_argument("--locale", default="pt-BR", choices=["pt-BR", "en-US"],
                    help="configuração regional do Access de origem (Format, CStr, CDate)")
    ap.add_argument("--incluir-ocultas", action="store_true",
                    help="inclui consultas ~sq_ (fontes de formulários/caixas de combinação)")
    ap.add_argument("--pacote", help="nome do pacote Python gerado (padrão: nome do arquivo)")
    ap.add_argument("--listar", action="store_true", help="só lista consultas, módulos e tabelas")
    args = ap.parse_args(argv)

    if args.tela:
        from .tela import servir
        return servir(origem=args.origem, saida=args.saida)
    if not args.origem:
        ap.error("informe o arquivo .accdb/.mdb (ou use --tela)")
    if args.validar:
        from .validacao import validar
        destino = args.resultado or "validacao.xlsx"
        resumo = validar(args.origem, args.validar, _parse_params(args.param), destino)
        ruins = resumo[resumo["Resultado"] != "IGUAL"] if len(resumo) else resumo
        print(f"Validação gravada em: {destino}")
        print("TUDO IGUAL ao Access." if len(resumo) and ruins.empty else
              f"{len(ruins)} item(ns) com diferença ou erro — veja o relatório.")
        return 0 if len(resumo) and ruins.empty else 3
    if args.listar:
        from .extract import open_extractor
        ex = open_extractor(args.origem, args.backend)
        qs = ex.queries(args.incluir_ocultas)
        print(f"Consultas ({len(qs)}):")
        for n in sorted(qs):
            print("  ", n)
        try:
            ms = ex.modules()
            print(f"Módulos VBA ({len(ms)}):")
            for m in ms:
                print("  ", m.name, f"({m.kind})")
        except Exception as e:  # noqa: BLE001
            print(f"Módulos VBA: não extraídos ({e})")
        ts = ex.table_names()
        print(f"Tabelas ({len(ts)}):")
        for t in sorted(ts):
            print("  ", t)
        return 0

    queries = list(args.consultas or [])
    if args.consultas_arquivo:
        with open(args.consultas_arquivo, encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.split("#", 1)[0].strip()
                if line:
                    queries.append(line)
    sequence = []
    if args.sequencia:
        with open(args.sequencia, encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.split("#", 1)[0].strip()
                if line:
                    sequence.append(line)
    res = transpile_database(args.origem, args.saida, queries=queries or None, sequence=sequence or None, modules=args.modulos,
                             no_vba=args.sem_vba, backend=args.backend,
                             export_data=args.exportar_dados, locale=args.locale,
                             include_hidden=args.incluir_ocultas, package=args.pacote,
                             input_modes={t: "mensal" for t in (args.mensais or [])},
                             all_monthly=args.todas_mensais)
    d = res.diags
    ok_q = sum(1 for _, s in res.queries.values() if s == "ok")
    ok_m = sum(1 for _, s in res.modules.values() if s.startswith("ok"))
    print(f"Pacote gerado em: {res.package_dir}")
    print(f"  consultas: {ok_q}/{len(res.queries)} convertidas"
          + (f" ({len(res.added_deps)} incluída(s) como dependência)" if res.added_deps else ""))
    if res.modules:
        print(f"  módulos VBA: {ok_m}/{len(res.modules)} convertidos")
    print(f"  avisos: {d.count('erro')} erro(s), {d.count('aviso')} aviso(s) — veja RELATORIO.md")
    if res.inputs:
        mens = [t for t, e in res.inputs.items() if e["modo"] == "mensal"]
        print(f"  entradas: {len(res.inputs)} tabela(s), {len(mens)} mensal(is): {', '.join(mens) or '-'}")
    if res.sequence:
        print(f"  sequência: {len(res.sequence)} etapa(s) em SEQUENCIA (consultas.py)")
    if args.resultado and res.sequence:
        pkg = _import_pkg(res.package_dir)
        print("Executando a sequência:")
        log = pkg.executar_sequencia(os.path.abspath(args.resultado), _parse_params(args.param),
                                     intermediarias=args.etapas_intermediarias)
        ok = (log["Situação"] == "ok").all()
        print(("Saída final gravada em: " if ok else "Sequência interrompida; log gravado em: ")
              + args.resultado)
        return 0 if ok else 2
    if args.resultado:
        resumo = _run_export(res.package_dir, args.resultado, _parse_params(args.param))
        falhas = resumo[resumo["Situação"] != "ok"]
        print(f"Resultado gravado em: {args.resultado} ({len(resumo) - len(falhas)}/{len(resumo)} consultas ok)")
        for _, r in falhas.iterrows():
            print(f"  ✗ {r['Consulta']}: {r['Erro']}")
    return 1 if d.count("erro") and ok_q == 0 and ok_m == 0 else 0


def _parse_params(items):
    import re
    from datetime import datetime
    out = {}
    for it in items:
        if "=" not in it:
            raise SystemExit(f"--param inválido (use NOME=VALOR): {it}")
        k, v = it.split("=", 1)
        v = v.strip()
        val = v
        if re.fullmatch(r"-?\d+", v):
            val = int(v)
        elif re.fullmatch(r"-?\d+[.,]\d+", v):
            val = float(v.replace(",", "."))
        else:
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
                try:
                    val = datetime.strptime(v, fmt)
                    break
                except ValueError:
                    pass
        out[k.strip()] = val
    return out


def _import_pkg(package_dir):
    import importlib
    sys.path.insert(0, str(package_dir.parent))
    return importlib.import_module(package_dir.name)


def _run_export(package_dir, destino, params):
    return _import_pkg(package_dir).exportar(os.path.abspath(destino), parametros=params)


if __name__ == "__main__":
    sys.exit(main())
