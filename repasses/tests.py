# -*- coding: utf-8 -*-
"""Suíte de regressão do motor de repasses.

Trava o comportamento validado (golds, casamento, classificação, OMIE) ANTES de
qualquer refatoração/otimização. Os testes de cálculo são SimpleTestCase (sem
banco): montam o livro a partir da planilha de regras e processam amostras reais.

Rodar:  python manage.py test repasses
"""

import io
import os
import unittest

from django.conf import settings
from django.test import SimpleTestCase

from repasses.services import medplus, omie, regras, repasse

_AMOSTRAS = r'C:\RepassesmedicosOMIE\amostras'
F18 = os.path.join(_AMOSTRAS, 'medplus_18_06.xls')             # Heric, com Valor (gold 4653.11)
FEXPORT = os.path.join(_AMOSTRAS, 'medplus_export.xls')        # 3 médicos (Tharcila gold 3250)
FFILIAIS = os.path.join(_AMOSTRAS, 'medplus_maio_filiais.xls')  # com coluna Clínica
PLANILHA = getattr(settings, 'REGRAS_REPASSE_PATH', '')


def _arquivos_presentes(*paths):
    return all(p and os.path.exists(p) for p in paths)


def _total_medico(rel, frag):
    """Soma dos honorários calculados (>0) do médico cujo nome contém frag."""
    total = 0.0
    for b in rel.blocos:
        if frag in b.profissional.lower():
            for p in b.procedimentos:
                if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                    total += p.honorario
    return round(total, 2)


@unittest.skipUnless(_arquivos_presentes(PLANILHA), 'planilha de regras ausente')
class MotorRegrasTests(SimpleTestCase):
    """Casamento de convênio/procedimento/nome e classificação de valor."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.livro = regras.carregar_regras(PLANILHA)

    def test_mapear_convenio(self):
        casos = {
            'OCI - SUS': 'oci', 'Cisamusep': 'cisa', 'Sus Maringá': 'sus',
            'PG Glaucoma': 'sus', 'Bradesco': 'particular', 'Amil': 'particular',
            'Particular': 'particular', 'Parcerias I': 'particular',
        }
        for conv, esperado in casos.items():
            self.assertEqual(regras.mapear_convenio(conv), esperado, conv)

    def test_classificar_valor_vazio_e_traco(self):
        self.assertEqual(regras._classificar_valor(''), (regras.SEM_VALOR, None))
        self.assertEqual(regras._classificar_valor('-'), (regras.NEGATIVO, 0.0))

    def test_valor_db_roundtrip(self):
        self.assertAlmostEqual(regras._valor_db('28.5%'), 0.285)   # ponto decimal em %
        self.assertAlmostEqual(regras._valor_db('28,5%'), 0.285)
        self.assertAlmostEqual(regras._valor_db('24%'), 0.24)
        self.assertEqual(regras._valor_db('1.120'), 1120.0)        # ponto milhar em fixo
        self.assertAlmostEqual(regras._valor_db('549,99'), 549.99)
        self.assertEqual(regras._valor_db('-'), '-')
        self.assertIsNone(regras._valor_db(''))

    def test_medico_por_nome(self):
        flavia = self.livro.medico_por_nome('Flávia Zandonadi')   # abreviação casa
        self.assertIsNotNone(flavia)
        self.assertIn('flavia', regras.normalizar(flavia.nome))
        mm = self.livro.medico_por_nome('Maria Marta Cabral')     # não casa nome alheio
        self.assertFalse(mm and 'marilia' in regras.normalizar(mm.nome))
        th = self.livro.medico_por_nome('Dra. Tharcila Breginski da Rocha (PR2)')  # sufixo
        self.assertIsNotNone(th)

    def test_calcular_casos_conhecidos(self):
        def calc(p, c, v=None, m='Dr. Heric Sakamoto'):
            return regras.calcular(self.livro, p, c, v, m)
        self.assertEqual(calc('Consulta', 'Sus Maringá').honorario, 25.0)
        self.assertEqual(calc('Consulta com Tonometria', 'Sus Maringá').honorario, 10.0)
        self.assertEqual(calc('Vitrectomia Posterior', 'Sus Maringá').honorario, 1000.0)
        self.assertEqual(calc('Vitrectomia Posterior com Infusão de Perfluocarbono',
                              'Sus Maringá').honorario, 1120.0)
        self.assertEqual(calc('Laudo - Retino/Angio', 'Sus Maringá').honorario, 10.0)
        self.assertEqual(calc('Laudo - Angio/Campi/Retino', 'Cisamusep').honorario, 30.0)
        self.assertAlmostEqual(calc('Sutura de Conjuntiva', 'Particular', 100.0).honorario, 24.0)

    def test_faco_consulta_cisa_usa_base_cisa(self):
        # decisão da diretoria: CISA = base CISA (150) + consulta (25) = 175
        proc = ('FACOEMULSIFICACAO C/ IMPLANTE DE LENTE INTRA-OCULAR DOBRAVEL '
                '(Incluso: Pre-Operatorio: Consulta de Avaliacao)')
        r = regras.calcular(self.livro, proc, 'Cisamusep', None, 'Dr. Heric Sakamoto')
        self.assertEqual(r.honorario, 175.0)


@unittest.skipUnless(_arquivos_presentes(PLANILHA, F18, FEXPORT), 'amostras ausentes')
class GoldsTests(SimpleTestCase):
    """Golds de ponta-a-ponta (parser + motor), via planilha (sem banco)."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.livro = regras.carregar_regras(PLANILHA)

    def _processado(self, arquivo):
        rel = medplus.ler_relatorio(arquivo)
        regras.processar(rel, self.livro)
        return rel

    def test_gold_heric_18_06(self):
        self.assertEqual(_total_medico(self._processado(F18), 'heric'), 4653.11)

    def test_gold_tharcila_export(self):
        self.assertEqual(_total_medico(self._processado(FEXPORT), 'tharcila'), 3250.0)


@unittest.skipUnless(_arquivos_presentes(PLANILHA, F18), 'amostras ausentes')
class RepasseDocTests(SimpleTestCase):
    """Total e reconciliação (arredondamento) do documento de repasse."""

    def setUp(self):
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(F18)
        regras.processar(rel, livro)
        self.bloco = next(b for b in rel.blocos if 'heric' in b.profissional.lower())

    def test_total_e_ajuste(self):
        total = repasse._total(self.bloco)
        soma_exibida = round(sum(round(p.honorario, 2) for p in repasse.pagaveis(self.bloco)), 2)
        self.assertEqual(total, 4653.11)
        # linhas exibidas + ajuste = total (documento de conferência fecha)
        self.assertEqual(round(soma_exibida + repasse._ajuste_arredondamento(self.bloco), 2), total)


@unittest.skipUnless(_arquivos_presentes(PLANILHA, FFILIAIS), 'amostra de filiais ausente')
class OmiePagarTests(SimpleTestCase):
    """Estrutura do contas a pagar (sem linha de categoria vazia)."""

    def setUp(self):
        from openpyxl import load_workbook
        from repasses import views
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(FFILIAIS)
        views._filtrar_blocos(rel)
        views._separar_por_dia(rel)
        regras.processar(rel, livro)
        views._limpar_linhas(rel)
        rel.anestesistas = []
        conteudo = omie.gerar_contas_pagar(rel, settings.OMIE_PAGAR_TEMPLATE).conteudo
        wb = load_workbook(io.BytesIO(conteudo))
        self.ws = wb[wb.sheetnames[0]]

    def test_sem_categoria_vazia(self):
        r, vazias, linhas = omie.LINHA_INICIAL, 0, 0
        while self.ws.cell(r, omie.COL_NOME).value:
            linhas += 1
            if not self.ws.cell(r, omie.COL_CATEGORIA).value:
                vazias += 1
            r += 1
        self.assertGreater(linhas, 0)
        self.assertEqual(vazias, 0, 'nenhuma linha do a pagar pode sair sem categoria')


class ParserUnitTests(SimpleTestCase):
    """Funções puras do parser (sem arquivos)."""

    def test_para_numero(self):
        self.assertAlmostEqual(medplus._para_numero('1.020,60'), 1020.6)
        self.assertAlmostEqual(medplus._para_numero('1020.6'), 1020.6)
        self.assertAlmostEqual(medplus._para_numero('25'), 25.0)
        self.assertIsNone(medplus._para_numero(''))
        self.assertIsNone(medplus._para_numero(None))

    def test_eh_cirurgia(self):
        self.assertTrue(medplus.eh_cirurgia('Facoemulsificação + LIO'))
        self.assertTrue(medplus.eh_cirurgia('Vitrectomia Posterior'))
        self.assertFalse(medplus.eh_cirurgia('Capsulotomia a YAG Laser'))
        self.assertFalse(medplus.eh_cirurgia('Consulta em Oftalmologia'))

    def test_subclasse_preview(self):
        def p(nome, classe):
            return medplus.Procedimento(None, '', '', nome, '', 1, None, None, classe=classe)
        self.assertEqual(medplus.subclasse_preview(p('Facoemulsificação', medplus.CLASSE_CIRURGIA)),
                         medplus.SUBCLASSE_CIRURGIA)
        self.assertEqual(medplus.subclasse_preview(p('YAG Laser', medplus.CLASSE_CIRURGIA)),
                         medplus.SUBCLASSE_PROCEDIMENTO)
        self.assertEqual(medplus.subclasse_preview(p('Consulta', medplus.CLASSE_EXAME)),
                         medplus.SUBCLASSE_EXAME)
