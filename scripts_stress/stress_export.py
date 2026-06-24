# -*- coding: utf-8 -*-
"""Teste de ESTRESSE de EXPORTACAO.

Processa a amostra real, replica os blocos para chegar a N blocos (200, 1000, ...)
e gera TODOS os arquivos de saida (OMIE a pagar/receber + Excel/PDF por bloco) e o
zip. Mede tempo, memoria (tracemalloc + RSS) e tamanho do zip. Banco isolado.
"""
import copy
import gc
import io
import os
import sys
import time
import tracemalloc
import zipfile

# --- bootstrap django ---------------------------------------------------------
BASE = r"C:\RepassesmedicosOMIE"
sys.path.insert(0, BASE)
os.environ["DJANGO_SETTINGS_MODULE"] = "config.settings"
import django  # noqa: E402
django.setup()

import shutil, tempfile  # noqa: E402
from django.conf import settings  # noqa: E402
from django.db import connection  # noqa: E402
_t = tempfile.mktemp(suffix=".sqlite3")
shutil.copyfile(connection.settings_dict["NAME"], _t)
connection.close()
connection.settings_dict["NAME"] = _t

from repasses.services import medplus, omie, regras, repasse  # noqa: E402
from repasses import views  # noqa: E402

AMOSTRA = r"C:\RepassesmedicosOMIE\amostras\medplus_22_06.xls"

try:
    import psutil  # noqa: E402
    _PROC = psutil.Process()
    def rss_mb():
        return _PROC.memory_info().rss / 1024 / 1024
except Exception:
    def rss_mb():
        return float("nan")


def processar_amostra():
    """Replica o pipeline de revisao (sem rascunho/edicoes) -> ResultadoImportacao."""
    resultado, aviso = views._ler_e_processar(AMOSTRA, "stress")
    resultado.log_edicoes = []
    views._aplicar_keiti(resultado)
    views._marcar_preceptoria(resultado, regras.carregar_livro_padrao())
    views._marcar_taxas_sala(resultado)
    views._limpar_linhas(resultado)
    views._indexar(resultado)
    return resultado, aviso


def _clonar_bloco(bloco, sufixo):
    """Copia profunda do bloco mudando o nome do profissional (arquivos unicos)."""
    novo = copy.deepcopy(bloco)
    novo.profissional = f"{bloco.profissional} #{sufixo}"
    return novo


def ampliar(resultado, n_alvo):
    """Replica os blocos da amostra ate ter >= n_alvo blocos PAGAVEIS.

    So conta blocos que geram repasse (pagaveis), pois sao os que viram Excel+PDF.
    """
    base = [b for b in resultado.blocos if repasse.pagaveis(b)]
    if not base:
        raise RuntimeError("amostra nao tem blocos pagaveis")
    base_anest = list(resultado.anestesistas)
    novos_blocos, novos_anest = [], []
    i = 0
    while len(novos_blocos) < n_alvo:
        for b in base:
            novos_blocos.append(_clonar_bloco(b, i))
            if len(novos_blocos) >= n_alvo:
                break
        for a in base_anest:
            na = copy.deepcopy(a)
            na["anestesista"] = f"{a.get('anestesista', '')} #{i}"
            novos_anest.append(na)
        i += 1
    res = medplus.ResultadoImportacao(unidade=resultado.unidade)
    res.blocos = novos_blocos
    res.anestesistas = novos_anest
    views._indexar(res)
    return res


def gerar_tudo(resultado):
    """Replica views._gerar_arquivos_por_dia + zip. Devolve (arquivos, zip_bytes, stats)."""
    t = {}
    t0 = time.perf_counter()
    pagar = omie.gerar_contas_pagar(resultado, settings.OMIE_PAGAR_TEMPLATE)
    t["pagar"] = time.perf_counter() - t0

    t0 = time.perf_counter()
    receber = omie.gerar_contas_receber(resultado, settings.OMIE_RECEBER_TEMPLATE)
    t["receber"] = time.perf_counter() - t0

    arquivos = [
        ("Importacao OMIE", "OMIE_Contas_a_Pagar.xlsx", pagar.conteudo),
        ("Importacao OMIE", "OMIE_Contas_a_Receber.xlsx", receber.conteudo),
    ]

    t0 = time.perf_counter()
    n_xlsx = n_pdf = 0
    for bloco in resultado.blocos:
        if not repasse.pagaveis(bloco):
            continue
        base = repasse.nome_base(bloco)
        grupo = f"Repasse - {bloco.profissional}"
        arquivos.append((grupo, f"{base}.xlsx", repasse.gerar_excel(bloco, resultado.unidade)))
        n_xlsx += 1
        arquivos.append((grupo, f"{base}.pdf", repasse.gerar_pdf(bloco, resultado.unidade)))
        n_pdf += 1
    t["repasses_medicos"] = time.perf_counter() - t0

    t0 = time.perf_counter()
    for a in resultado.anestesistas:
        base = repasse.nome_base_anestesista(a)
        grupo = f"Anestesista - {a['anestesista']}"
        arquivos.append((grupo, f"{base}.xlsx", repasse.gerar_excel_anestesista(a, resultado.unidade)))
        arquivos.append((grupo, f"{base}.pdf", repasse.gerar_pdf_anestesista(a, resultado.unidade)))
    t["repasses_anest"] = time.perf_counter() - t0

    # zip (igual ao baixar_lote_zip: zipa os bytes)
    t0 = time.perf_counter()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        usados = {}
        for (g, n, conteudo) in arquivos:
            pasta = (g or "").strip().replace("/", "-") or "Arquivos"
            nome = f"{pasta}/{n}"
            usados[nome] = usados.get(nome, 0) + 1
            if usados[nome] > 1:
                b, _, ext = n.rpartition(".")
                nome = f"{pasta}/{b} ({usados[nome]}).{ext}" if ext else f"{pasta}/{n} ({usados[nome]})"
            zf.writestr(nome, bytes(conteudo))
    zip_bytes = buf.getvalue()
    t["zip"] = time.perf_counter() - t0

    bruto = sum(len(c) for (_, _, c) in arquivos)
    stats = dict(t, n_arquivos=len(arquivos), n_xlsx=n_xlsx, n_pdf=n_pdf,
                 n_linhas_pagar=pagar.linhas, n_linhas_receber=receber.linhas,
                 bytes_brutos=bruto, bytes_zip=len(zip_bytes))
    return arquivos, zip_bytes, stats


def fmt_mb(b):
    return f"{b/1024/1024:.1f} MB"


TRACE = os.environ.get("STRESS_TRACE", "1") == "1"


def rodar(resultado_base, n_alvo, aviso_template=None):
    res = ampliar(resultado_base, n_alvo)
    gc.collect()
    if TRACE:
        tracemalloc.start()
    rss0 = rss_mb()
    t0 = time.perf_counter()
    try:
        arquivos, zip_bytes, stats = gerar_tudo(res)
        erro = None
    except Exception as exc:
        import traceback
        erro = f"{type(exc).__name__}: {exc}"
        traceback.print_exc()
        stats = {}
    total = time.perf_counter() - t0
    if TRACE:
        cur, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()
    else:
        peak = 0.0
    rss1 = rss_mb()
    print(f"\n=== {n_alvo} blocos pagaveis ===")
    if erro:
        print("  ERRO:", erro)
        return {"n_alvo": n_alvo, "erro": erro}
    print(f"  arquivos gerados : {stats['n_arquivos']} "
          f"({stats['n_xlsx']} xlsx + {stats['n_pdf']} pdf + 2 OMIE + anest)")
    print(f"  linhas OMIE      : pagar={stats['n_linhas_pagar']} receber={stats['n_linhas_receber']}")
    print(f"  tempo total      : {total:.1f}s")
    print(f"    a pagar        : {stats['pagar']:.2f}s")
    print(f"    a receber      : {stats['receber']:.2f}s")
    print(f"    repasses med   : {stats['repasses_medicos']:.2f}s "
          f"({stats['repasses_medicos']/max(stats['n_xlsx'],1)*1000:.1f} ms/bloco xlsx+pdf)")
    print(f"    anestesistas   : {stats['repasses_anest']:.2f}s")
    print(f"    zip            : {stats['zip']:.2f}s")
    print(f"  bytes brutos     : {fmt_mb(stats['bytes_brutos'])}")
    print(f"  tamanho zip      : {fmt_mb(stats['bytes_zip'])} "
          f"(compressao {stats['bytes_zip']/max(stats['bytes_brutos'],1)*100:.0f}%)")
    print(f"  tracemalloc peak : {fmt_mb(peak)}")
    print(f"  RSS              : {rss0:.0f} -> {rss1:.0f} MB (delta {rss1-rss0:+.0f})")
    return dict(stats, n_alvo=n_alvo, total_s=total, peak_mb=peak/1024/1024,
                rss_delta=rss1-rss0)


def main():
    alvos = [int(x) for x in sys.argv[1:]] or [200, 1000]
    print("Lendo e processando a amostra real...")
    t0 = time.perf_counter()
    resultado, aviso = processar_amostra()
    dt = time.perf_counter() - t0
    n_pag = sum(1 for b in resultado.blocos if repasse.pagaveis(b))
    print(f"  amostra: {len(resultado.blocos)} blocos ({n_pag} pagaveis), "
          f"{len(resultado.anestesistas)} anestesistas, "
          f"{resultado.total_registros} registros, lida em {dt:.1f}s")
    if aviso:
        print("  AVISO:", aviso)
    resumo = []
    for n in alvos:
        resumo.append(rodar(resultado, n))
    print("\n\n=== RESUMO ===")
    print(f"{'blocos':>8} {'arquivos':>9} {'tempo(s)':>9} {'zip':>10} "
          f"{'peak':>9} {'RSS d':>8} {'ms/bloco':>9}")
    for r in resumo:
        if r.get("erro"):
            print(f"{r['n_alvo']:>8}  ERRO: {r['erro']}")
            continue
        msb = r["repasses_medicos"] / max(r["n_xlsx"], 1) * 1000
        print(f"{r['n_alvo']:>8} {r['n_arquivos']:>9} {r['total_s']:>9.1f} "
              f"{fmt_mb(r['bytes_zip']):>10} {fmt_mb(r['peak_mb']*1024*1024):>9} "
              f"{r['rss_delta']:>+7.0f} {msb:>9.1f}")


if __name__ == "__main__":
    main()
