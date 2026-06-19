# -*- coding: utf-8 -*-
"""Correções memorizadas: aplicação e gravação.

Quando a diretoria ajusta um honorário na revisão e marca "memorizar", o valor
é guardado (models.CorrecaoMemorizada) e passa a ser reaplicado automaticamente
toda vez que o mesmo procedimento/convênio aparecer — a correção "chega sozinha"
ao próximo processamento, sem ninguém refazer na mão.
"""

from __future__ import annotations

import re

from ..models import CorrecaoMemorizada
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


def aplicar(resultado) -> int:
    """Sobrepõe os honorários do resultado com as correções memorizadas ativas.

    Retorna quantas linhas foram corrigidas. Mantém a precisão cheia (o
    arredondamento para Reais é só na exibição/escrita final)."""
    ativos = list(CorrecaoMemorizada.objects.filter(ativo=True))
    if not ativos:
        return 0
    index = {(c.proc_norm, c.conv_norm, c.medico_norm): c for c in ativos}

    n = 0
    for bloco in resultado.blocos:
        med_n = normalizar(_sem_sufixo(bloco.profissional))
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
