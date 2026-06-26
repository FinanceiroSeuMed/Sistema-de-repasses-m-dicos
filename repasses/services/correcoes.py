# -*- coding: utf-8 -*-
"""Correções memorizadas: aplicação e gravação.

Quando a diretoria ajusta um honorário na revisão e marca "memorizar", o valor
é guardado (models.CorrecaoMemorizada) e passa a ser reaplicado automaticamente
toda vez que o mesmo procedimento/convênio aparecer — a correção "chega sozinha"
ao próximo processamento, sem ninguém refazer na mão.
"""

from __future__ import annotations

import re

from ..models import ClasseMemorizada, CorrecaoMemorizada
from .regras import normalizar

_RE_SUFIXO = re.compile(r'\s*\([^)]*\)\s*$')

# Status que NÃO são sobrepostos pela memória (resolvidos em fluxo próprio).
_NAO_SOBRESCREVER = {'catarata', 'componente'}


def _sem_sufixo(nome: str) -> str:
    return _RE_SUFIXO.sub('', nome or '').strip()


def _buscar(index, proc_n, conv_n, med_n):
    """Procura a correção mais específica para (proc, convênio, médico).

    Ordem de preferência: médico+convênio > médico (qualquer conv) >
    convênio (qualquer médico) > genérica (qualquer médico e convênio)."""
    for chave in ((proc_n, conv_n, med_n),
                  (proc_n, '', med_n),
                  (proc_n, conv_n, ''),
                  (proc_n, '', '')):
        c = index.get(chave)
        if c is not None:
            return c
    return None


def medico_norm(profissional, livro=None) -> str:
    """Chave normalizada do médico, RESOLVIDA pelo cadastro quando possível.

    O nome no MedPlus ("Dra. Tharcila Breginski da Rocha (PR2)") quase nunca é igual
    ao do cadastro ("Dra. Tharcila..."). Para a correção memorizada de um médico casar
    de verdade nos próximos meses, os dois lados precisam usar a MESMA chave — então,
    havendo livro, casamos o nome pelo matcher por tokens (livro.medico_por_nome) e
    usamos o nome do CADASTRO como chave canônica. Sem livro/sem casar: o próprio nome."""
    if livro is not None:
        m = livro.medico_por_nome(profissional)
        if m is not None:
            return normalizar(_sem_sufixo(m.nome))
    return normalizar(_sem_sufixo(profissional))


def aplicar(resultado, livro=None) -> int:
    """Sobrepõe os honorários do resultado com as correções memorizadas ativas.

    Retorna quantas linhas foram corrigidas. Mantém a precisão cheia (o
    arredondamento para Reais é só na exibição/escrita final). `livro` (opcional)
    resolve o médico pelo cadastro para a correção por-médico casar (ver medico_norm)."""
    ativos = list(CorrecaoMemorizada.objects.filter(ativo=True))
    if not ativos:
        return 0
    index = {(c.proc_norm, c.conv_norm, c.medico_norm): c for c in ativos}

    n = 0
    for bloco in resultado.blocos:
        med_n = medico_norm(bloco.profissional, livro)
        for p in bloco.procedimentos:
            if p.status_calculo in _NAO_SOBRESCREVER:
                continue
            c = _buscar(index, normalizar(p.procedimento), normalizar(p.convenio), med_n)
            if c is None:
                continue
            if c.tipo == CorrecaoMemorizada.TIPO_PERCENTUAL:
                if p.valor is None:
                    continue  # percentual precisa do valor bruto
                p.honorario = float(c.valor) * float(p.valor)
            else:
                p.honorario = float(c.valor)
            if c.classe:
                p.classe = c.classe  # também lembra a mudança de classe (ex.: laudos)
            p.status_calculo = 'calculado'
            p.motivo_calculo = f'Correção memorizada (salva em {c.criado_em:%d/%m/%Y}).'
            n += 1
    return n


def aplicar_classes(resultado) -> int:
    """Reaplica a classe memorizada de cada procedimento (vale para TODOS os médicos).

    A classificação é intrínseca ao procedimento — então, diferente do honorário,
    casa só pelo nome do procedimento (qualquer médico/convênio). Roda no
    processamento, ANTES da revisão, para o procedimento já vir classificado."""
    ativos = list(ClasseMemorizada.objects.filter(ativo=True))
    if not ativos:
        return 0
    index = {c.proc_norm: c.classe for c in ativos}
    n = 0
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            # Taxa de sala é definida pelo CONVÊNIO (não pelo nome): a memória por
            # procedimento não pode reescrever a classe dessas linhas, senão o
            # _marcar_taxas_sala deixa de reconhecê-las. (Diretoria 2026-06-26.)
            if (p.status_calculo in _NAO_SOBRESCREVER
                    or 'taxa' in normalizar(p.convenio)):
                continue
            classe = index.get(normalizar(p.procedimento))
            if classe and p.classe != classe:
                p.classe = classe
                n += 1
    return n


def memorizar_classe(procedimento, classe, *, origem=''):
    """Cria/atualiza a classe memorizada de um procedimento (upsert por nome)."""
    obj, _criado = ClasseMemorizada.objects.update_or_create(
        proc_norm=normalizar(procedimento),
        defaults={'procedimento': procedimento, 'classe': classe,
                  'ativo': True, 'origem': origem},
    )
    return obj


def memorizar(procedimento, convenio, valor, *, medico='', tipo=CorrecaoMemorizada.TIPO_FIXO,
              classe='', origem='', observacao=''):
    """Cria ou atualiza uma correção. Casa pelos campos normalizados (upsert)."""
    proc_n = normalizar(procedimento)
    conv_n = normalizar(convenio)
    med_n = normalizar(_sem_sufixo(medico))
    obj, _criado = CorrecaoMemorizada.objects.update_or_create(
        proc_norm=proc_n, conv_norm=conv_n, medico_norm=med_n,
        defaults={
            'procedimento': procedimento, 'convenio': convenio or '',
            'medico': medico or '', 'tipo': tipo, 'valor': valor, 'classe': classe or '',
            'ativo': True, 'origem': origem, 'observacao': observacao,
        },
    )
    return obj
