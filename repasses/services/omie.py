# -*- coding: utf-8 -*-
"""
Geradores dos arquivos de importação da OMIE (contas a pagar e a receber)
a partir de um relatório da MedPlus já processado (com honorários calculados).

Regras definidas pela diretoria:
- Conta corrente: "Omie.CASH".
- Vencimento: dia 10 do mês seguinte ao mês do atendimento.
- A pagar: uma linha somada por (médico × classe), com a categoria da classe.
  Só entram honorários CALCULADOS (> 0). Linhas "a definir" não entram (ficam
  pendentes para o usuário preencher).
- A receber: uma linha somada por convênio (Cliente = nome do convênio), com o
  valor bruto pago pelo paciente/convênio.
"""

from __future__ import annotations

import io
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date

from openpyxl import load_workbook

from . import medplus

# Mapeamento classe -> categoria da OMIE (contas a pagar)
CATEGORIA_POR_CLASSE = {
    medplus.CLASSE_CIRURGIA: 'Repasse Oftalmologistas - Cirurgia',
    medplus.CLASSE_EXAME: 'Repasse Oftalmologistas - Exame',
    medplus.CLASSE_PRECEPTORIA: 'Preceptoria',
    # Anestesistas: 'Repasse Anestesiologistas' (quando houver essa classe)
}

CONTA_CORRENTE = 'Omie.CASH'

# Colunas (1-indexadas) no layout das planilhas OMIE; dados começam na linha 6.
LINHA_INICIAL = 6
COL_NOME = 3        # C: Fornecedor (pagar) / Cliente (receber)
COL_CATEGORIA = 4   # D
COL_CONTA = 5       # E
COL_VALOR = 6       # F
COL_REGISTRO = 10   # J: Data de Registro
COL_VENCIMENTO = 11 # K: Data de Vencimento


@dataclass
class ResultadoSaida:
    nome_arquivo: str
    conteudo: bytes
    linhas: int
    pendencias: list = field(default_factory=list)


def venc_dia10_mes_seguinte(d: date) -> date:
    if d.month == 12:
        return date(d.year + 1, 1, 10)
    return date(d.year, d.month + 1, 10)


def _data_referencia(resultado) -> date:
    datas = [p.data for b in resultado.blocos for p in b.procedimentos if p.data]
    return max(datas) if datas else date.today()


def _fmt(d: date) -> str:
    return d.strftime('%d/%m/%Y')


def _escrever(modelo_path, linhas: list[dict]) -> tuple[bytes, int]:
    """linhas: lista de dicts com chaves nome, categoria, valor, registro, vencimento."""
    wb = load_workbook(modelo_path)
    ws = wb[wb.sheetnames[0]]
    r = LINHA_INICIAL
    for ln in linhas:
        ws.cell(row=r, column=COL_NOME, value=ln['nome'])
        ws.cell(row=r, column=COL_CATEGORIA, value=ln['categoria'])
        ws.cell(row=r, column=COL_CONTA, value=CONTA_CORRENTE)
        ws.cell(row=r, column=COL_VALOR, value=round(ln['valor'], 2))
        ws.cell(row=r, column=COL_REGISTRO, value=_fmt(ln['registro']))
        ws.cell(row=r, column=COL_VENCIMENTO, value=_fmt(ln['vencimento']))
        r += 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), len(linhas)


def gerar_contas_pagar(resultado, modelo_path) -> ResultadoSaida:
    ref = _data_referencia(resultado)
    venc = venc_dia10_mes_seguinte(ref)
    pendencias = []
    grupos = defaultdict(float)        # (medico, razao, classe) -> soma honorario
    for bloco in resultado.blocos:
        nome_forn = bloco.razao_social or bloco.profissional
        if not bloco.razao_social:
            pendencias.append(f'{bloco.profissional}: sem Razão Social (usado o nome) — confira na OMIE.')
        for p in bloco.procedimentos:
            if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                grupos[(nome_forn, p.classe)] += p.honorario
            elif p.status_calculo == 'a_definir':
                pendencias.append(f'{bloco.profissional}: "{p.procedimento[:40]}" a definir — não entrou no a pagar.')

    linhas = []
    for (nome, classe), soma in grupos.items():
        categoria = CATEGORIA_POR_CLASSE.get(classe, '')
        if not categoria:
            pendencias.append(f'Classe "{classe}" sem categoria OMIE definida — linha de {nome} ficou sem categoria.')
        linhas.append({'nome': nome, 'categoria': categoria, 'valor': soma,
                       'registro': ref, 'vencimento': venc})
    linhas.sort(key=lambda x: (x['nome'], x['categoria']))
    conteudo, n = _escrever(modelo_path, linhas)
    return ResultadoSaida('OMIE_Contas_a_Pagar.xlsx', conteudo, n, pendencias)


def gerar_contas_receber(resultado, modelo_path, categoria_receber: str) -> ResultadoSaida:
    ref = _data_referencia(resultado)
    venc = venc_dia10_mes_seguinte(ref)
    pendencias = []
    grupos = defaultdict(float)        # convenio -> soma valor bruto
    sem_valor = 0
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            if p.valor is None:
                sem_valor += 1
                continue
            convenio = p.convenio or '(sem convênio)'
            grupos[convenio] += p.valor
    if sem_valor:
        pendencias.append(f'{sem_valor} procedimento(s) sem valor bruto (arquivo do médico não traz o valor) — '
                          'use o relatório completo da MedPlus para o contas a receber.')

    linhas = []
    for convenio, soma in sorted(grupos.items()):
        linhas.append({'nome': convenio, 'categoria': categoria_receber, 'valor': soma,
                       'registro': ref, 'vencimento': venc})
    conteudo, n = _escrever(modelo_path, linhas)
    return ResultadoSaida('OMIE_Contas_a_Receber.xlsx', conteudo, n, pendencias)
