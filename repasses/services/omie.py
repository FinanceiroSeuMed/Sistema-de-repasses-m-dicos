# -*- coding: utf-8 -*-
"""
Geradores dos arquivos de importação da OMIE (contas a pagar e a receber)
a partir de um relatório da MedPlus já processado (com honorários calculados).

Regras definidas pela diretoria:
- Conta corrente: "Omie.CASH".
- Vencimento: dia 10 do mês seguinte ao mês do atendimento.
- A pagar: uma linha por (médico × dia × clínica × classe), categoria da classe.
  Só entram honorários CALCULADOS (> 0). Linhas "a definir" ou sem categoria não
  entram (viram pendência para o usuário resolver na revisão).
- A receber: uma linha por (dia × clínica) — recebimento geral dos pacientes —
  com o valor bruto somado; Cliente = a associação (CLIENTE_RECEBER), filial no
  Departamento; categoria = fallback único.
"""

from __future__ import annotations

import io
import re
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
}
CATEGORIA_ANESTESISTA = 'Repasses Anestesiologistas'

# Taxa de sala NÃO entra no a pagar (só no a receber).
CLASSES_FORA_DO_PAGAR = {medplus.CLASSE_TAXA}

CONTA_CORRENTE = 'Omie.CASH'

# Contas a RECEBER: o Cliente somos nós (a associação); a filial vai no Departamento.
# Na OMIE os clientes são todos "ASSOCIACAO SEUMED HOSPITAL DE OLHOS" por filial.
CLIENTE_RECEBER = 'ASSOCIACAO SEUMED HOSPITAL DE OLHOS'

# Colunas (1-indexadas) no layout das planilhas OMIE; dados começam na linha 6.
LINHA_INICIAL = 6
COL_NOME = 3        # C: Fornecedor (pagar) / Cliente (receber)
COL_CATEGORIA = 4   # D
COL_CONTA = 5       # E
COL_VALOR = 6       # F
COL_REGISTRO = 10   # J: Data de Registro
COL_VENCIMENTO = 11 # K: Data de Vencimento
COL_OBSERVACOES = 19            # S: Observações (igual nos dois modelos)
COL_DEPARTAMENTO_PAGAR = 50     # Departamento (100%) — modelo a pagar
COL_DEPARTAMENTO_RECEBER = 42   # Departamento (100%) — modelo a receber

_RE_SUFIXO = re.compile(r'\s*\([^)]*\)\s*$')


def _sem_sufixo(nome: str) -> str:
    """Remove o sufixo de unidade entre parênteses do nome do profissional.

    "Dr. Carlos Eduardo (PR2)" -> "Dr. Carlos Eduardo" (a clínica já vai no
    Departamento; na observação fica só o nome do médico).
    """
    return _RE_SUFIXO.sub('', nome or '').strip()


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


def _escrever(modelo_path, linhas: list[dict], col_departamento: int) -> tuple[bytes, int]:
    """linhas: dicts com nome, categoria, valor, registro, vencimento, observacao, departamento.

    Preserva a formatação do template oficial: datas vão como VALOR de data real
    (não texto) com formato dd/mm/aaaa, e o valor como número 18,2 — para a OMIE
    importar sem erro de formato. A coluna de Departamento difere entre os dois
    modelos (a pagar e a receber), por isso vem como parâmetro.
    """
    wb = load_workbook(modelo_path)
    ws = wb[wb.sheetnames[0]]
    r = LINHA_INICIAL
    for ln in linhas:
        ws.cell(row=r, column=COL_NOME, value=ln['nome'])
        ws.cell(row=r, column=COL_CATEGORIA, value=ln['categoria'])
        ws.cell(row=r, column=COL_CONTA, value=CONTA_CORRENTE)
        cval = ws.cell(row=r, column=COL_VALOR, value=round(float(ln['valor']), 2))
        cval.number_format = '0.00'
        for col, chave in ((COL_REGISTRO, 'registro'), (COL_VENCIMENTO, 'vencimento')):
            cel = ws.cell(row=r, column=col, value=ln[chave])  # objeto date, não string
            cel.number_format = 'DD/MM/YYYY'
        if ln.get('observacao'):
            ws.cell(row=r, column=COL_OBSERVACOES, value=ln['observacao'])
        if ln.get('departamento'):
            ws.cell(row=r, column=col_departamento, value=ln['departamento'])
        r += 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), len(linhas)


def gerar_contas_pagar(resultado, modelo_path) -> ResultadoSaida:
    """Uma linha por (médico × dia × clínica × classe).

    Cada bloco já representa um médico em um dia numa clínica; dentro do bloco
    somamos por classe (cirurgia/exame/preceptoria). Se o Dr. atende em duas
    clínicas em dois dias, saem quatro repasses — quatro linhas no a pagar.
    Observação = "Repasse dd/mm Nome do Dr."; Departamento = nome da clínica.
    """
    ref = _data_referencia(resultado)
    pendencias = []
    linhas = []
    for bloco in resultado.blocos:
        nome_forn = bloco.razao_social or bloco.profissional
        if not bloco.razao_social:
            pendencias.append(f'{bloco.profissional}: sem Razão Social (usado o nome) — confira na OMIE.')
        dia = bloco.data or ref
        venc = venc_dia10_mes_seguinte(dia)
        medico = _sem_sufixo(bloco.profissional)
        observacao = f'Repasse {dia.strftime("%d/%m")} {medico}'
        por_classe = defaultdict(float)    # classe -> soma honorário (precisão cheia)
        for p in bloco.procedimentos:
            if p.classe in CLASSES_FORA_DO_PAGAR:
                continue  # taxa de sala só vai no a receber
            if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                por_classe[p.classe] += p.honorario
            elif p.status_calculo == 'a_definir':
                pendencias.append(f'{bloco.profissional}: "{p.procedimento[:40]}" a definir — não entrou no a pagar.')
        for classe, soma in por_classe.items():
            categoria = CATEGORIA_POR_CLASSE.get(classe, '')
            if not categoria:
                # Sem categoria OMIE (ex.: "A classificar") NÃO entra no arquivo —
                # a OMIE rejeita linha sem categoria. Vira pendência para a pessoa
                # reclassificar na revisão antes de exportar.
                pendencias.append(f'{nome_forn}: R$ {soma:.2f} em "{classe}" — classifique '
                                  f'(Cirurgia/Exame/Preceptoria) na revisão; NÃO entrou no a pagar.')
                continue
            linhas.append({'nome': nome_forn, 'categoria': categoria, 'valor': soma,
                           'registro': dia, 'vencimento': venc,
                           'observacao': observacao, 'departamento': bloco.clinica})

    # Linhas dos anestesistas (categoria própria) — uma por atendimento/dia
    for a in getattr(resultado, 'anestesistas', []):
        dia = a.get('data') or ref
        anest = _sem_sufixo(a.get('anestesista', ''))
        linhas.append({'nome': a.get('razao_social') or a.get('anestesista'),
                       'categoria': CATEGORIA_ANESTESISTA, 'valor': a.get('valor', 0),
                       'registro': dia, 'vencimento': venc_dia10_mes_seguinte(dia),
                       'observacao': f'Repasse {dia.strftime("%d/%m")} {anest}',
                       'departamento': a.get('clinica', '')})

    linhas.sort(key=lambda x: (str(x['registro']), x['departamento'] or '', x['nome'], x['categoria']))
    conteudo, n = _escrever(modelo_path, linhas, COL_DEPARTAMENTO_PAGAR)
    return ResultadoSaida('OMIE_Contas_a_Pagar.xlsx', conteudo, n, pendencias)


def gerar_contas_receber(resultado, modelo_path, categoria_fallback='Outras Receitas com Serviços') -> ResultadoSaida:
    """Uma linha por (dia × clínica) — recebimento geral dos pacientes.

    Soma o valor bruto de todos os atendimentos daquele dia naquela clínica.
    Observação = "Recebimento de atendimentos dd/mm Clínica"; Departamento = clínica.
    """
    ref = _data_referencia(resultado)
    pendencias = []
    grupos = defaultdict(float)        # (dia, clínica) -> soma valor bruto
    sem_valor = 0
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            if p.valor is None:
                sem_valor += 1
                continue
            grupos[(bloco.data, bloco.clinica or '')] += p.valor
    if sem_valor:
        pendencias.append(f'{sem_valor} procedimento(s) sem valor bruto (arquivo do médico não traz o valor) — '
                          'use o relatório completo da MedPlus para o contas a receber.')

    linhas = []
    for (dia, clinica), soma in sorted(grupos.items(), key=lambda kv: (str(kv[0][0]), kv[0][1])):
        d = dia or ref
        rotulo = f' {clinica}' if clinica else ''
        linhas.append({'nome': CLIENTE_RECEBER, 'categoria': categoria_fallback,
                       'valor': soma, 'registro': d, 'vencimento': venc_dia10_mes_seguinte(d),
                       'observacao': f'Recebimento de atendimentos {d.strftime("%d/%m")}{rotulo}',
                       'departamento': clinica})
    conteudo, n = _escrever(modelo_path, linhas, COL_DEPARTAMENTO_RECEBER)
    return ResultadoSaida('OMIE_Contas_a_Receber.xlsx', conteudo, n, pendencias)
