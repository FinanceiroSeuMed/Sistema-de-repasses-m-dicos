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


# Razão social + CNPJ por médico (diretoria 2026-06-28). O CNPJ é a CHAVE do
# Fornecedor no contas a pagar da OMIE. Casa por trecho do nome normalizado.
_PJ_OVERRIDES = {
    'heric':      ('HERIC MASSAAKI SAKAMOTO', '18.145.426/0001-07'),
    'tiezzi':     ('TIEZZE & TIEZZI SERVIÇOS MEDICOS LTDA', '65.358.679/0001-30'),
    'zandonadi':  ('ZANDONADI SERVIÇOS MEDICOS LTDA', '39.898.771/0001-83'),
    'alessander': ('ALESSANDER TIEO TSUNETO LTDA', '34.952.198/0001-25'),
    'carolina':   ('DUARTE CARDIOLOGIA LTDA', '30.848.132/0001-39'),
    'tharcila':   ('GASTREN - CLINICA MEDICA S/S LTDA', '17.357.063/0001-00'),
    'thalia':     ('THALIA MACARIS OFTALMOLOGIA LTDA', '59.883.866/0001-30'),
    'isabela':    ('ISABELA MIWA MAEDA CLINICA MEDICA', '35.773.858/0001-73'),
    'marilia':    ('M A F PAPA SERVIÇOS MEDICOS LTDA', '39.776.675/0001-62'),
    'suellen':    ('S.H.O SERVICOS MEDICOS LTDA', '41.050.531/0001-76'),
    'keiti':      ('CLINICA SHIRASU LTDA', '15.240.760/0001-43'),
    'rodolpho':   ('IKIGAI OFTALMOLOGIA LTDA', '58.456.887/0001-07'),
    'ana paula':  ('GUERMANDI SERVICOS MEDICOS LTDA', '22.295.697/0001-08'),
    'licia':      ('LICIA INAZAWA DA SILVA CLINICA MEDICA LTDA', '35.736.953/0001-05'),
    'roberta':    ('R.S.GOMES LTDA', '22.743.146/0001-60'),
}

# Observações padronizadas pela diretoria (2026-06-27).
_OBS_RESIDENTE = 'Não recebem em Oftalmo'
_OBS_ANEST_CASSIANA = 'R$ 1.200,00 + R$ 200,00 por hora extra (mesma regra da Dra. Cassiana).'


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
            # override da diretoria (tem prioridade sobre a planilha): razão + CNPJ.
            # O CNPJ é a chave do Fornecedor no contas a pagar da OMIE.
            for frag, (razao, cnpj) in _PJ_OVERRIDES.items():
                if frag in chave:
                    medico.razao_social = razao
                    medico.cnpj = cnpj
            if rm.obs:
                medico.regra_obs = rm.obs
            # Observações padronizadas (diretoria 2026-06-27):
            #  - TODOS os residentes: "Não recebem em Oftalmo".
            #  - Anestesistas Suellen/Isabela/Marília: mesma regra da Dra. Cassiana.
            if cat == Medico.CATEGORIA_RESIDENTE:
                medico.regra_obs = _OBS_RESIDENTE
            if medico.eh_anestesista and any(t in chave for t in ('suellen', 'isabela', 'marilia')):
                medico.regra_obs = _OBS_ANEST_CASSIANA
            # Dr. Keiti é preceptor da "Equipe Dr. Keiti" — aparece na lista de responsáveis
            # da agenda dele. (Diretoria 2026-07-01.)
            if 'keiti' in chave:
                medico.eh_preceptor = True
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
