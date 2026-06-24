# -*- coding: utf-8 -*-
"""Context processors globais (disponíveis em todos os templates)."""


def edicao_em_andamento(request):
    """Sinaliza no menu que há um repasse em edição (rascunho do último import).
    Usa a MESMA checagem do continuar_edicao (rascunho com arquivo resolvível), p/ o
    link do menu e o continuar não divergirem (link não fica preso se o upload sumiu)."""
    try:
        from .views import _rascunho_em_andamento
        tem = _rascunho_em_andamento() is not None
    except Exception:
        tem = False
    return {'tem_edicao': tem}
