# -*- coding: utf-8 -*-
"""
Leitor das regras de repasse (anexo 5) e motor de cálculo do honorário.

Princípios (definidos pela diretoria):
- Casamento APROXIMADO de procedimentos (abreviações, sinônimos). Ex.: catarata =
  facoemulsificação/facectomia.
- É melhor PERGUNTAR e deixar em branco do que arriscar um valor errado. Honorário
  R$ 0,00 só quando a regra diz isso ("-", "não recebem").
- Percentuais (ex.: 0,24) incidem sobre o VALOR pago pelo paciente/convênio.

A planilha tem muitas exceções; este motor resolve com segurança os casos diretos
(consulta/exame/procedimento com valor fixo ou percentual) e marca os demais como
"a definir" para o usuário preencher.
"""

from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass, field
from difflib import SequenceMatcher

import pandas as pd

# Status do cálculo de uma linha
CALCULADO = 'calculado'      # encontrou regra e calculou
NAO_RECEBE = 'nao_recebe'    # regra diz que não há repasse (0,00)
A_DEFINIR = 'a_definir'      # sem regra clara -> em branco para o usuário

# Tipos de valor de uma regra
FIXO = 'fixo'
PERCENTUAL = 'percentual'
MANUAL = 'manual'            # texto/observação -> exige decisão humana
SEM_VALOR = 'sem_valor'      # célula vazia para aquele convênio
NEGATIVO = 'negativo'        # "-" -> não recebe

_PAGADORES = ('particular', 'convenio', 'sus', 'oci', 'cisa')

# palavras pouco informativas, ignoradas no casamento
_STOP = {'a', 'o', 'de', 'da', 'do', 'com', 'e', 'c', 'em', 'por', '-', 'ao', 'mono'}

# injeção de sinônimos: se o procedimento contém a chave, adiciona o token
_SINONIMOS = {
    'facoemulsificacao': 'catarata',
    'facectomia': 'catarata',
    'coerencia': 'oct',
}


def normalizar(texto) -> str:
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ''
    t = str(texto).strip().lower()
    t = ''.join(c for c in unicodedata.normalize('NFKD', t) if not unicodedata.combining(c))
    t = re.sub(r'[^a-z0-9 ]+', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


def _tokens(texto: str) -> set[str]:
    # remove o conteúdo entre parênteses ANTES de tokenizar — ex.: o trecho
    # "(Incluso: ... Consulta de Avaliação)" não deve fazer uma cirurgia casar
    # com a regra "Consulta".
    sem_parenteses = re.sub(r'\([^)]*\)', ' ', str(texto or ''))
    sem_parenteses = re.sub(r'\(.*$', ' ', sem_parenteses)  # parêntese que não fecha (texto truncado)
    base = normalizar(sem_parenteses)
    toks = {p for p in base.split() if p not in _STOP and len(p) > 1}
    for chave, syn in _SINONIMOS.items():
        if chave in base:
            toks.add(syn)
    return toks


def _classificar_valor(bruto):
    """Interpreta uma célula de preço da planilha de regras."""
    if bruto is None or (isinstance(bruto, float) and pd.isna(bruto)):
        return SEM_VALOR, None
    if isinstance(bruto, (int, float)):
        v = float(bruto)
        if v <= 0:
            return NEGATIVO, 0.0
        return (PERCENTUAL, v) if v < 1 else (FIXO, v)
    texto = str(bruto).strip()
    if texto in ('-', '–', ''):
        return NEGATIVO, 0.0
    # número em texto?
    limpo = texto.replace('R$', '').replace('.', '').replace(',', '.').strip()
    try:
        v = float(limpo)
        if v <= 0:
            return NEGATIVO, 0.0
        return (PERCENTUAL, v) if v < 1 else (FIXO, v)
    except ValueError:
        return MANUAL, texto  # ex.: "Ver", "mesmo valor da catarata"


# --- Estruturas ---------------------------------------------------------------

@dataclass
class RegraProcedimento:
    classe: str
    nome: str
    nome_norm: str
    tokens: set
    valores: dict           # pagador -> célula bruta


@dataclass
class Medico:
    nome: str
    categoria: str
    razao_social: str = ''
    obs: str = ''


@dataclass
class LivroRegras:
    procedimentos: list[RegraProcedimento] = field(default_factory=list)
    medicos: list[Medico] = field(default_factory=list)
    lembretes_preceptoria: list[str] = field(default_factory=list)

    def medico_por_nome(self, nome: str) -> Medico | None:
        alvo = normalizar(nome)
        melhor, score = None, 0.0
        for m in self.medicos:
            s = SequenceMatcher(None, alvo, normalizar(m.nome)).ratio()
            if s > score:
                melhor, score = m, s
        return melhor if score >= 0.6 else None


# --- Carregamento da planilha de regras --------------------------------------

def _achar_cabecalho_pagadores(raw):
    for i in range(len(raw)):
        mapa = {}
        for j, cel in enumerate(raw.iloc[i]):
            n = normalizar(cel)
            if n in _PAGADORES:
                mapa[n] = j
        if 'particular' in mapa and ('sus' in mapa or 'convenio' in mapa):
            return i, mapa
    return None, {}


def _ler_aba_precos(caminho, aba, classe) -> list[RegraProcedimento]:
    raw = pd.read_excel(caminho, sheet_name=aba, engine='openpyxl', header=None)
    lin_cab, pagadores = _achar_cabecalho_pagadores(raw)
    if lin_cab is None:
        return []
    primeira_pag = min(pagadores.values())
    regras = []
    for i in range(lin_cab + 1, len(raw)):
        linha = raw.iloc[i]
        # nome = célula de texto mais à direita antes das colunas de pagador
        nome = ''
        for j in range(primeira_pag):
            cel = linha.iloc[j] if j < len(linha) else None
            t = '' if cel is None or (isinstance(cel, float) and pd.isna(cel)) else str(cel).strip()
            if t:
                nome = t
        if not nome:
            continue
        valores = {}
        for pag, col in pagadores.items():
            valores[pag] = linha.iloc[col] if col < len(linha) else None
        if all(_classificar_valor(v)[0] == SEM_VALOR for v in valores.values()):
            continue  # linha sem nenhum preço (cabeçalho de seção)
        regras.append(RegraProcedimento(
            classe=classe, nome=nome, nome_norm=normalizar(nome), tokens=_tokens(nome), valores=valores,
        ))
    return regras


def _ler_medicos(caminho) -> tuple[list[Medico], list[str]]:
    raw = pd.read_excel(caminho, sheet_name='Médicos', engine='openpyxl', header=None)
    medicos, lembretes = [], []
    categoria = ''

    def _cel(linha, j):
        return str(linha.iloc[j]).strip() if linha.shape[0] > j and pd.notna(linha.iloc[j]) else ''

    for i in range(len(raw)):
        linha = raw.iloc[i]
        c1, c2, c3, c4 = _cel(linha, 1), _cel(linha, 2), _cel(linha, 3), _cel(linha, 4)
        if c2.lower() == 'médicos':   # linha de cabeçalho da aba
            continue
        if c1:                        # rótulo de categoria (pode estar na mesma linha do 1º médico)
            categoria = c1
        if c2:
            medicos.append(Medico(nome=c2, categoria=categoria, razao_social=c4, obs=c3))
            if 'preceptor' in categoria.lower() and c3:
                lembretes.append(f'{c2}: {c3}')
    return medicos, lembretes


def carregar_regras(caminho) -> LivroRegras:
    livro = LivroRegras()
    livro.procedimentos += _ler_aba_precos(caminho, 'Consultas e Exames', 'Exames e Consultas')
    livro.procedimentos += _ler_aba_precos(caminho, 'Cirurgias e Procedimentos', 'Cirurgias e Procedimentos')
    livro.medicos, livro.lembretes_preceptoria = _ler_medicos(caminho)
    return livro


# --- Cálculo ------------------------------------------------------------------

def mapear_convenio(convenio: str) -> str | None:
    n = normalizar(convenio)
    if not n:
        return None
    # ordem importa: "OCI - SUS" é OCI (não SUS); CISA antes de SUS
    if 'oci' in n:
        return 'oci'
    if 'cisa' in n:
        return 'cisa'
    if 'sus' in n or n.startswith('pg'):
        return 'sus'
    if 'particular' in n:
        return 'particular'
    if 'conven' in n:
        return 'convenio'
    return None


@dataclass
class ResultadoCalculo:
    status: str
    honorario: float | None
    motivo: str = ''
    regra: str = ''
    tipo: str = ''


def _similaridade(tokens_proc: set, regra: RegraProcedimento) -> tuple[float, int]:
    """Quão bem as palavras da regra estão CONTIDAS no procedimento.

    Só contém (não usa similaridade de sequência, que gerava falsos positivos
    como 'recobrimento conjuntival' casar com 'tumor conjuntival'). Devolve
    (proporção de palavras da regra presentes, nº de palavras casadas).
    """
    if not regra.tokens:
        return 0.0, 0
    casadas = len(regra.tokens & tokens_proc)
    return casadas / len(regra.tokens), casadas


def calcular(livro: LivroRegras, procedimento: str, convenio: str, valor, medico: str = '',
             limiar: float = 0.6) -> ResultadoCalculo:
    pagador = mapear_convenio(convenio)
    if pagador is None:
        return ResultadoCalculo(A_DEFINIR, None, motivo=f'Convênio não reconhecido: "{convenio}".')

    # Regra por categoria do médico
    m = livro.medico_por_nome(medico) if medico else None
    proc_norm = normalizar(procedimento)
    eh_catarata = ('catarata' in _tokens(procedimento))
    if m and 'residente' in m.categoria.lower():
        return ResultadoCalculo(NAO_RECEBE, 0.0, motivo='Residente — não recebe honorário.')
    if m and 'fellow' in m.categoria.lower() and eh_catarata:
        return ResultadoCalculo(NAO_RECEBE, 0.0, motivo='Fellow — não recebe em catarata.')

    tokens_proc = _tokens(procedimento)
    candidatos = []
    for regra in livro.procedimentos:
        tipo, _ = _classificar_valor(regra.valores.get(pagador))
        if tipo == SEM_VALOR:
            continue
        proporcao, casadas = _similaridade(tokens_proc, regra)
        if proporcao >= limiar:
            candidatos.append((proporcao, casadas, regra))
    if not candidatos:
        return ResultadoCalculo(A_DEFINIR, None,
                                motivo='Sem regra correspondente — preencher manualmente.')
    # melhor proporção; empate -> regra que casou mais palavras (mais específica)
    candidatos.sort(key=lambda x: (x[0], x[1]), reverse=True)
    _, _, regra = candidatos[0]
    bruto = regra.valores.get(pagador)
    tipo, val = _classificar_valor(bruto)

    if tipo == NEGATIVO:
        return ResultadoCalculo(NAO_RECEBE, 0.0, motivo=f'Regra "{regra.nome}" não prevê repasse neste convênio.',
                                regra=regra.nome, tipo=tipo)
    if tipo == MANUAL:
        return ResultadoCalculo(A_DEFINIR, None, motivo=f'Regra especial: "{val}".', regra=regra.nome, tipo=tipo)
    if tipo == PERCENTUAL:
        if valor is None:
            return ResultadoCalculo(A_DEFINIR, None, motivo='Percentual sem valor bruto informado.',
                                    regra=regra.nome, tipo=tipo)
        return ResultadoCalculo(CALCULADO, round(float(valor) * val, 2),
                                motivo=f'{val*100:.0f}% de {valor} (regra "{regra.nome}").',
                                regra=regra.nome, tipo=tipo)
    # FIXO
    return ResultadoCalculo(CALCULADO, round(val, 2),
                            motivo=f'Valor fixo (regra "{regra.nome}").', regra=regra.nome, tipo=tipo)


# --- Orquestração: aplicar regras a um relatório lido -------------------------

def carregar_livro_padrao() -> LivroRegras | None:
    """Carrega as regras do caminho configurado em settings.REGRAS_REPASSE_PATH."""
    from django.conf import settings
    caminho = getattr(settings, 'REGRAS_REPASSE_PATH', '')
    if caminho and os.path.exists(caminho):
        return carregar_regras(caminho)
    return None


def processar(resultado, livro: LivroRegras):
    """Preenche honorário/status de cada procedimento e os lembretes por médico."""
    for bloco in resultado.blocos:
        medico = livro.medico_por_nome(bloco.profissional)
        if medico:
            bloco.razao_social = medico.razao_social
            if 'preceptor' in medico.categoria.lower() and medico.obs:
                bloco.lembrete = f'Repasse de preceptoria a lançar à parte: {medico.obs}'
        for p in bloco.procedimentos:
            r = calcular(livro, p.procedimento, p.convenio, p.valor, bloco.profissional)
            p.honorario = r.honorario
            p.status_calculo = r.status
            p.motivo_calculo = r.motivo
    return resultado
