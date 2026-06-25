# -*- coding: utf-8 -*-
"""Relatório mensal compilado — junta os repasses (a pagar) do mês num único xlsx,
ordenado por médico, no formato "Repasses em Aberto" (anexo da diretoria):
colunas Filial | Destino | DataVencimento | Valor | Categoria, e um resumo por
classe (Consultas e exames / Cirurgias e procedimentos / Preceptoria / TOTAL).
"""

from __future__ import annotations

import io
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

_ORDEM_RESUMO = ['Consultas e exames', 'Cirurgias e procedimentos', 'Preceptoria', 'Anestesia']
_MESES = ['', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho',
          'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']


def nome_mes(aaaa_mm: str) -> str:
    """'2026-05' -> 'Maio 2026'."""
    try:
        ano, mes = aaaa_mm.split('-')
        return f'{_MESES[int(mes)]} {ano}'
    except Exception:
        return aaaa_mm


def _data(iso: str):
    try:
        return date.fromisoformat(iso)
    except Exception:
        return None


def gerar_relatorio_mensal(linhas: list[dict], titulo: str = 'Repasses em Aberto') -> bytes:
    """Gera o xlsx do mês. `linhas` = dicts de omie.linhas_relatorio_pagar."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Contas a Pagar - Padrão'

    negrito = Font(bold=True)
    cab_fill = PatternFill('solid', fgColor='D9E1F2')
    esq = Alignment(horizontal='left')

    for c, t in enumerate(['Filial', 'Destino', 'DataVencimento', 'Valor', 'Categoria'], start=1):
        cel = ws.cell(1, c, t)
        cel.font = negrito
        cel.fill = cab_fill

    # Dados ordenados por médico (Destino) e depois por data — "separado por Dr.".
    ls = sorted(linhas, key=lambda x: ((x.get('medico') or '').lower(), x.get('data') or ''))
    r = 2
    for ln in ls:
        ws.cell(r, 1, ln.get('departamento') or ln.get('clinica') or '')
        ws.cell(r, 2, ln.get('medico') or '')
        cel_v = ws.cell(r, 3, _data(ln.get('vencimento')))
        cel_v.number_format = 'DD/MM/YYYY'
        cel_val = ws.cell(r, 4, round(float(ln.get('valor') or 0), 2))
        cel_val.number_format = '#,##0.00'
        ws.cell(r, 5, ln.get('categoria') or '')
        r += 1

    # Resumo por classe (colunas G/H), na ordem padrão + TOTAL.
    resumo = defaultdict(float)
    for ln in linhas:
        resumo[ln.get('resumo') or 'Outros'] += float(ln.get('valor') or 0)
    presentes = [k for k in _ORDEM_RESUMO if k in resumo] + \
                [k for k in resumo if k not in _ORDEM_RESUMO]
    rr = 1
    for k in presentes:
        ws.cell(rr, 7, k).font = negrito
        cel = ws.cell(rr, 8, round(resumo[k], 2))
        cel.number_format = '#,##0.00'
        cel.alignment = esq
        rr += 1
    ws.cell(rr, 7, 'TOTAL').font = negrito
    cel_t = ws.cell(rr, 8, round(sum(resumo.values()), 2))
    cel_t.font = negrito
    cel_t.number_format = '#,##0.00'
    cel_t.alignment = esq

    for col, w in zip('ABCDE', (18, 32, 16, 12, 34)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions['G'].width = 26
    ws.column_dimensions['H'].width = 13

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
