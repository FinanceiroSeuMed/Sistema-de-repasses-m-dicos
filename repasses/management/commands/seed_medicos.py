# -*- coding: utf-8 -*-
"""Popula/atualiza o cadastro de médicos a partir da planilha de regras (anexo 5)."""

import re

from django.core.management.base import BaseCommand

from repasses.models import Medico
from repasses.services import regras


def _corrigir_nome(nome: str) -> str:
    """Corrige erros de digitação herdados da planilha de regras.
    Ex.: "Redientes Dra. Regina" -> "Residentes Dra. Regina". (Diretoria 2026-06-26.)"""
    return re.sub(r'(?i)rediente', 'Residente', nome)   # "Redientes" -> "Residentes"


# Razões sociais confirmadas pela diretoria que NÃO vêm na planilha de regras.
# Casa por trecho do nome normalizado do médico (ex.: "keiti").
_RAZAO_OVERRIDES = {
    'keiti': 'CLINICA SHIRASU',
}


def _categoria_primaria(cat_norm: str) -> str:
    if 'residente' in cat_norm:
        return Medico.CATEGORIA_RESIDENTE
    if 'fellow' in cat_norm:
        return Medico.CATEGORIA_FELLOW
    if 'preceptor' in cat_norm:
        return Medico.CATEGORIA_PRECEPTOR
    if 'anestesista' in cat_norm:
        return Medico.CATEGORIA_ANESTESISTA
    return Medico.CATEGORIA_MEDICO


class Command(BaseCommand):
    help = 'Popula/atualiza o cadastro de médicos a partir da planilha de regras (anexo 5).'

    def handle(self, *args, **options):
        livro = regras.carregar_livro_padrao()
        if livro is None:
            self.stderr.write(self.style.ERROR(
                'Planilha de regras não encontrada (settings.REGRAS_REPASSE_PATH).'))
            return

        # Corrige nomes com erro de digitação que já estejam no cadastro (ex.: "Redientes")
        # ANTES de montar o índice, para o seed casar com o registro renomeado (não duplicar).
        for m in Medico.objects.all():
            novo = _corrigir_nome(m.nome)
            if novo != m.nome and not Medico.objects.filter(nome=novo).exclude(pk=m.pk).exists():
                m.nome = novo
                m.save(update_fields=['nome'])

        existentes = {regras.normalizar(m.nome): m for m in Medico.objects.all()}
        criados = atualizados = 0

        for rm in livro.medicos:
            nome = _corrigir_nome(rm.nome)
            chave = regras.normalizar(nome)
            if not chave:
                continue
            cat = _categoria_primaria(regras.normalizar(rm.categoria))

            medico = existentes.get(chave)
            if medico is None:
                medico = Medico(nome=nome)
                criados += 1
            else:
                atualizados += 1

            # categoria primária: não rebaixa um papel já definido para "médico"
            if medico.categoria in ('', Medico.CATEGORIA_MEDICO):
                medico.categoria = cat
            # acumula papéis (um médico pode ser fellow E preceptor)
            if cat == Medico.CATEGORIA_FELLOW:
                medico.eh_fellow = True
            if cat == Medico.CATEGORIA_PRECEPTOR:
                medico.eh_preceptor = True
            if cat == Medico.CATEGORIA_ANESTESISTA:
                medico.eh_anestesista = True

            if rm.razao_social:
                medico.razao_social = rm.razao_social
            # override da diretoria (tem prioridade sobre a planilha)
            for frag, razao in _RAZAO_OVERRIDES.items():
                if frag in chave:
                    medico.razao_social = razao
            if rm.obs:
                medico.regra_obs = rm.obs
            medico.save()
            existentes[chave] = medico

        # Dra. Regina (indivíduo) — paga em DINHEIRO, fora da OMIE. A "Equipe/Residentes
        # Dra. Regina" da planilha recebe repasse normal. (Diretoria 2026-06-25.)
        Medico.objects.get_or_create(
            nome='Dra. Regina',
            defaults={'categoria': Medico.CATEGORIA_ANESTESISTA, 'eh_anestesista': True,
                      'regra_obs': 'Paga em dinheiro: < 24 cirurgias = R$ 1.000; ≥ 24 = R$ 1.500.'})

        self.stdout.write(self.style.SUCCESS(
            f'Cadastro de médicos atualizado: {criados} criados, {atualizados} atualizados '
            f'({Medico.objects.count()} no total).'))
