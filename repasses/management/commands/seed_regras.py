# -*- coding: utf-8 -*-
"""Importa as regras de honorário da planilha (anexo 5) para o banco.

Depois disso, as regras passam a ser geridas DENTRO do sistema (admin / página
de Regras) e a planilha deixa de ser necessária. Rodar uma vez:

    python manage.py seed_regras            # cria/atualiza pelas regras da planilha
    python manage.py seed_regras --limpar   # apaga as regras atuais antes
"""

from django.conf import settings
from django.core.management.base import BaseCommand

from repasses.models import RegraRepasse
from repasses.services import regras


class Command(BaseCommand):
    help = 'Importa as regras de honorário da planilha para o banco de dados.'

    def add_arguments(self, parser):
        parser.add_argument('--planilha', default=getattr(settings, 'REGRAS_REPASSE_PATH', ''))
        parser.add_argument('--limpar', action='store_true',
                            help='Apaga todas as regras antes de importar.')

    def handle(self, *args, **opts):
        caminho = opts['planilha']
        livro = regras.carregar_regras(caminho)
        if not livro.procedimentos:
            self.stderr.write(self.style.ERROR(f'Nenhuma regra lida de {caminho}.'))
            return

        if opts['limpar']:
            n = RegraRepasse.objects.count()
            RegraRepasse.objects.all().delete()
            self.stdout.write(f'{n} regra(s) antiga(s) apagada(s).')

        criadas = atualizadas = 0
        for rp in livro.procedimentos:
            db_vals = {f'val_{pag}': regras._celula_para_db(rp.valores.get(pag))
                       for pag in regras._PAGADORES}
            _, criada = RegraRepasse.objects.update_or_create(
                nome_norm=regras.normalizar(rp.nome), classe=rp.classe,
                defaults={'nome': rp.nome, **db_vals})
            criadas += 1 if criada else 0
            atualizadas += 0 if criada else 1

        self._vitrectomia_como_catarata()
        self.stdout.write(self.style.SUCCESS(
            f'OK: {criadas} criada(s), {atualizadas} atualizada(s). '
            f'Total no banco: {RegraRepasse.objects.count()}.'))

    def _vitrectomia_como_catarata(self):
        """A 'Vitrectomia Anterior' é tratada como catarata: 4 linhas espelhando as da
        catarata (mesmos valores). Substitui a linha antiga "mesmo valor da catarata".
        (Diretoria 2026-06-27.) Idempotente — sobrevive ao re-seed."""
        cataratas = list(RegraRepasse.objects.filter(nome_norm__startswith='cirurgia de catarata'))
        feitas = 0
        for cat in cataratas:
            novo_nome = cat.nome.replace('Cirurgia de Catarata', 'Vitrectomia Anterior')
            if novo_nome == cat.nome:
                continue
            RegraRepasse.objects.update_or_create(
                nome_norm=regras.normalizar(novo_nome), classe=cat.classe,
                defaults={'nome': novo_nome, 'val_particular': cat.val_particular,
                          'val_convenio': cat.val_convenio, 'val_sus': cat.val_sus,
                          'val_oci': cat.val_oci, 'val_cisa': cat.val_cisa, 'ativo': True})
            feitas += 1
        # remove a linha única antiga ("Vitrectomia Anterior" com texto não-numérico)
        RegraRepasse.objects.filter(nome_norm='vitrectomia anterior').delete()
        if feitas:
            self.stdout.write(f'Vitrectomia Anterior: {feitas} linha(s) espelhando a catarata.')
