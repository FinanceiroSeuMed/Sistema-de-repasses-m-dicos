# -*- coding: utf-8 -*-
"""Popula/atualiza o cadastro de médicos a partir da planilha de regras (anexo 5)."""

from django.core.management.base import BaseCommand

from repasses.models import Medico
from repasses.services import regras


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

        existentes = {regras.normalizar(m.nome): m for m in Medico.objects.all()}
        criados = atualizados = 0

        for rm in livro.medicos:
            chave = regras.normalizar(rm.nome)
            if not chave:
                continue
            cat = _categoria_primaria(regras.normalizar(rm.categoria))

            medico = existentes.get(chave)
            if medico is None:
                medico = Medico(nome=rm.nome)
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

        self.stdout.write(self.style.SUCCESS(
            f'Cadastro de médicos atualizado: {criados} criados, {atualizados} atualizados '
            f'({Medico.objects.count()} no total).'))
