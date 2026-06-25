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
from django.test import Client, SimpleTestCase, TestCase

from repasses.services import medplus, omie, regras, repasse

_AMOSTRAS = r'C:\RepassesmedicosOMIE\amostras'
F18 = os.path.join(_AMOSTRAS, 'medplus_18_06.xls')             # Heric, com Valor (gold 4653.11)
FEXPORT = os.path.join(_AMOSTRAS, 'medplus_export.xls')        # 3 médicos (Tharcila gold 3250)
FFILIAIS = os.path.join(_AMOSTRAS, 'medplus_maio_filiais.xls')  # com coluna Clínica
FEXPORT2 = os.path.join(_AMOSTRAS, 'medplus_export_2.xls')      # tem OCI de residentes (Vida/Arthur)
F22 = os.path.join(_AMOSTRAS, 'medplus_22_06.xls')             # Heric Qtd/tono (gold Matriz 2449.11, PR2/PR3 1220)
PLANILHA = getattr(settings, 'REGRAS_REPASSE_PATH', '')


def _arquivos_presentes(*paths):
    return all(p and os.path.exists(p) for p in paths)


def _total_medico(rel, frag, clinica=None):
    """Soma dos honorários calculados (>0) do médico cujo nome contém frag.
    Se `clinica` for dado, soma só os procedimentos daquela clínica."""
    total = 0.0
    for b in rel.blocos:
        if frag in b.profissional.lower():
            for p in b.procedimentos:
                if clinica is not None and p.clinica != clinica:
                    continue
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
        def calc(p, c, v=None, m='Dr. Heric Sakamoto', q=1):
            return regras.calcular(self.livro, p, c, v, m, quantidade=q)
        self.assertEqual(calc('Consulta', 'Sus Maringá').honorario, 25.0)
        # Consulta SUS: diferencia tonografia (paciente pagou < R$30 -> R$10) da
        # consulta de avaliação normal (>= R$30 ou sem valor -> R$25). O texto
        # "(Tono)" no nome NÃO decide — só o valor pago. (Diretoria 2026-06-23.)
        self.assertEqual(calc('Consulta - Avaliação (Tono(1))', 'Sus Maringá', 12.0).honorario, 10.0)
        self.assertEqual(calc('Consulta - Avaliação (Tono(1))', 'Sus Maringá', 38.37).honorario, 25.0)
        # Qtd multiplica o valor fixo (mapeamento binocular Qtd=2 -> dobra).
        self.assertEqual(calc('Mapeamento de Retina', 'Sus Maringá', 98.48, q=2).honorario, 50.0)
        self.assertEqual(calc('Mapeamento de Retina', 'Sus Maringá', 98.48).honorario, 25.0)
        self.assertEqual(calc('Vitrectomia Posterior', 'Sus Maringá').honorario, 1000.0)
        self.assertEqual(calc('Vitrectomia Posterior com Infusão de Perfluocarbono',
                              'Sus Maringá').honorario, 1120.0)
        self.assertEqual(calc('Laudo - Retino/Angio', 'Sus Maringá').honorario, 10.0)
        self.assertEqual(calc('Laudo - Angio/Campi/Retino', 'Cisamusep').honorario, 30.0)
        self.assertAlmostEqual(calc('Sutura de Conjuntiva', 'Particular', 100.0).honorario, 24.0)
        # Copel e Sanepar seguem a regra de particular (consulta = R$60, como o particular)
        part = calc('Consulta em Oftalmologia', 'Particular', 150.0).honorario
        self.assertEqual(calc('Consulta em Oftalmologia', 'Copel', 150.0).honorario, part)
        self.assertEqual(calc('Consulta em Oftalmologia', 'Sanepar', 150.0).honorario, part)
        self.assertEqual(regras.mapear_convenio('Copel'), 'particular')
        self.assertEqual(regras.mapear_convenio('Sanepar'), 'particular')

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

    @unittest.skipUnless(_arquivos_presentes(F22), 'amostra 22/06 ausente')
    def test_gold_heric_22_06(self):
        # Gold 22/06: Qtd (mapeamento binocular=2 -> R$50) e consulta-tono por valor
        # (avaliação SUS R$38,37 -> R$25, não R$10). Dois blocos por clínica.
        rel = self._processado(F22)
        self.assertEqual(_total_medico(rel, 'heric', 'Maringá - Matriz'), 2449.11)
        self.assertEqual(_total_medico(rel, 'heric', 'Maringá - Filial PR2 e PR3'), 1220.0)


@unittest.skipUnless(_arquivos_presentes(PLANILHA, F18), 'amostras ausentes')
class RepasseDocTests(SimpleTestCase):
    """Total e reconciliação (arredondamento) do documento de repasse."""

    def setUp(self):
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(F18)
        regras.processar(rel, livro)
        self.bloco = next(b for b in rel.blocos if 'heric' in b.profissional.lower())

    def test_linhas_somam_o_total(self):
        # Sem linha de "Arredondamento": a soma das linhas EXIBIDAS fecha com o Total
        # (a diferença de centavos foi embutida numa linha). Documento de conferência.
        cab, dados, total, n = repasse._linhas_repasse(self.bloco)
        self.assertEqual(total, 4653.11)
        soma_exibida = round(sum(linha[5] for linha in dados), 2)   # coluna Honorário
        self.assertEqual(soma_exibida, total)
        self.assertNotIn('Arredondamento', cab)


@unittest.skipUnless(_arquivos_presentes(PLANILHA, F22), 'amostra 22/06 ausente')
class RepasseLayoutTests(SimpleTestCase):
    """Excel e PDF do médico: mesmas colunas (com Paciente), célula de Total e SEM
    aviso de preceptoria nas exportações."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        from openpyxl import load_workbook
        from pypdf import PdfReader
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(F22)
        regras.processar(rel, livro)
        cls.bloco = next(b for b in rel.blocos if 'heric' in b.profissional.lower())
        cls.bloco.lembrete = 'Repasse de preceptoria a lançar à parte: teste'  # não pode vazar p/ export
        cls.wb = load_workbook(io.BytesIO(repasse.gerar_excel(cls.bloco, 'Maringá')))
        cls.ws = cls.wb.active
        cls.pdf_txt = '\n'.join((p.extract_text() or '')
                                for p in PdfReader(io.BytesIO(repasse.gerar_pdf(cls.bloco, 'Maringá'))).pages)

    def _linha_cab(self):
        return next(r for r in range(1, 15) if self.ws.cell(r, 1).value == 'Data')

    def test_excel_colunas_e_total(self):
        cr = self._linha_cab()
        self.assertEqual([self.ws.cell(cr, c).value for c in range(1, 7)],
                         ['Data', 'Paciente', 'Procedimento', 'Convênio', 'Qtd.', 'Honorário'])
        # Total à DIREITA: rótulo "Total" na col 5 (Qtd) e valor na col 6 (Honorário).
        tot = next(r for r in range(cr, self.ws.max_row + 1) if self.ws.cell(r, 5).value == 'Total')
        esperado = repasse._total(self.bloco)
        self.assertAlmostEqual(float(self.ws.cell(tot, 6).value), esperado)
        pac = next(p.paciente for p in repasse.pagaveis(self.bloco) if p.paciente)
        self.assertEqual(self.ws.cell(cr + 1, 2).value, pac)   # 1ª linha traz o paciente

    def test_pdf_igual_ao_excel(self):
        self.assertIn('Paciente', self.pdf_txt)        # mesma coluna do Excel
        self.assertIn('Total', self.pdf_txt)
        self.assertNotIn('preceptoria a lançar', self.pdf_txt.lower())  # aviso só na preview


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
        views._marcar_taxas_sala(rel)
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


class OmieReceberTests(SimpleTestCase):
    """De-para clínica->CNPJ e categoria por grupo de convênio (a receber)."""

    def test_cnpj_e_grupo(self):
        self.assertEqual(omie.cnpj_filial('Maringá - Matriz'), '27.717.567/0001-30')
        self.assertEqual(omie.cnpj_filial('Mandaguaçu'), '27.717.567/0002-10')
        self.assertEqual(omie.cnpj_filial('Maringá - Filial PR2 e PR3'), '27.717.567/0007-25')
        self.assertIsNone(omie.cnpj_filial('Clínica Inexistente'))
        self.assertEqual(omie.grupo_receber('Sus Maringá'), 'SUS')
        self.assertEqual(omie.grupo_receber('OCI - SUS'), 'SUS')
        self.assertEqual(omie.grupo_receber('Cisamusep'), 'CISAMUSEP')
        self.assertEqual(omie.grupo_receber('Particular'), 'PARTICULARES')
        self.assertEqual(omie.grupo_receber('Parcerias I'), 'PARTICULARES')

    def test_departamento_texto(self):
        # Texto "NN. Nome" (sem acento), mesmo no a pagar e no a receber.
        self.assertEqual(omie.departamento('Maringá - Matriz'), '01. Matriz')
        self.assertEqual(omie.departamento('Mandaguaçu'), '02. Mandaguacu')
        self.assertEqual(omie.departamento('Paiçandu'), '03. Paicandu')
        self.assertEqual(omie.departamento('Sarandi'), '04. Sarandi')
        self.assertEqual(omie.departamento('Maringá - Av Brasil'), '05. Brasil')
        self.assertEqual(omie.departamento('Mandaguari'), '06. Mandaguari')
        # PR2 e PR3 (mesmo CNPJ 07): glaucoma/PR3 -> "07. PR3"; senão "07. PR2".
        self.assertEqual(omie.departamento('Maringá - Filial PR2 e PR3', 'PR3'), '07. PR3')
        self.assertEqual(omie.departamento('Maringá - Filial PR2 e PR3', 'PR2'), '07. PR2')
        self.assertEqual(omie.departamento('Maringá - Filial PR2 e PR3'), '07. PR2')
        self.assertIsNone(omie.departamento('Outra'))

    def test_deteccao_pr3(self):
        from repasses import views
        self.assertTrue(views._eh_pr3('Dra. Fulana (PR3)'))
        self.assertTrue(views._eh_pr3('Agenda Glaucoma'))
        self.assertTrue(views._eh_pr3('Dr. Beltrano - Glaucoma'))
        self.assertFalse(views._eh_pr3('Dra. Fulana (PR2)'))
        self.assertFalse(views._eh_pr3('Dr. Heric Sakamoto'))

    @unittest.skipUnless(_arquivos_presentes(PLANILHA, F22), 'amostra 22/06 ausente')
    def test_receber_22_06_estrutura(self):
        from openpyxl import load_workbook
        from repasses import views
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(F22)
        views._filtrar_blocos(rel); views._separar_por_dia(rel)
        regras.processar(rel, livro)
        rel.blocos = [b for b in rel.blocos if not regras.eh_residente(livro, b.profissional)]
        views._marcar_taxas_sala(rel); views._limpar_linhas(rel)
        res = omie.gerar_contas_receber(rel, settings.OMIE_RECEBER_TEMPLATE)
        wb = load_workbook(io.BytesIO(res.conteudo)); ws = wb[wb.sheetnames[0]]
        r, linhas = omie.LINHA_INICIAL, 0
        while ws.cell(r, omie.COL_NOME).value:
            linhas += 1
            cli = ws.cell(r, omie.COL_NOME).value
            cat = ws.cell(r, omie.COL_CATEGORIA).value
            self.assertRegex(str(cli), r'^\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}$')   # Cliente = CNPJ
            self.assertIn(cat, ('SUS', 'CISAMUSEP', 'PARTICULARES'))
            self.assertEqual(ws.cell(r, omie.COL_CONTA).value, omie.CONTA_CORRENTE)
            r += 1
        self.assertGreater(linhas, 0)


@unittest.skipUnless(_arquivos_presentes(PLANILHA, F22), 'amostra 22/06 ausente')
class RelatorioMensalTests(SimpleTestCase):
    """Relatório mensal compilado por Dr. (formato Repasses em Aberto)."""

    def test_formato_e_resumo(self):
        from openpyxl import load_workbook
        from repasses.services import relatorio
        from repasses import views
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(F22)
        views._filtrar_blocos(rel); views._separar_por_dia(rel)
        regras.processar(rel, livro)
        rel.blocos = [b for b in rel.blocos if not regras.eh_residente(livro, b.profissional)]
        views._marcar_taxas_sala(rel); views._limpar_linhas(rel)

        linhas = omie.linhas_relatorio_pagar(rel)
        self.assertTrue(linhas)
        for ln in linhas:
            self.assertRegex(ln['data'], r'^\d{4}-\d{2}-\d{2}$')          # data ISO (p/ JSON)
            self.assertIn(ln['resumo'], ('Consultas e exames', 'Cirurgias e procedimentos',
                                         'Preceptoria', 'Anestesia'))
        conteudo = relatorio.gerar_relatorio_mensal(linhas, 'Teste')
        ws = load_workbook(io.BytesIO(conteudo))['Contas a Pagar - Padrão']
        self.assertEqual([ws.cell(1, c).value for c in range(1, 6)],
                         ['Filial', 'Destino', 'DataVencimento', 'Valor', 'Categoria'])
        nomes = [ws.cell(r, 2).value for r in range(2, 2 + len(linhas))]
        self.assertEqual(nomes, sorted(nomes, key=lambda s: s.lower()))   # ordenado por Dr.
        r, classes = 1, []
        while ws.cell(r, 7).value and ws.cell(r, 7).value != 'TOTAL':
            classes.append(ws.cell(r, 8).value); r += 1
        self.assertEqual(ws.cell(r, 7).value, 'TOTAL')
        self.assertAlmostEqual(round(sum(classes), 2), round(ws.cell(r, 8).value, 2))


@unittest.skipUnless(_arquivos_presentes(PLANILHA, FEXPORT2), 'amostra com OCI de residente ausente')
class OciResidentesTests(SimpleTestCase):
    """OCI feito por residente vai para o repasse do Dr. Alessander."""

    def test_oci_residente_vai_para_alessander(self):
        from repasses import views
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(FEXPORT2)
        views._filtrar_blocos(rel)
        views._separar_por_dia(rel)
        regras.processar(rel, livro)
        raw_res = sum(1 for b in rel.blocos if regras.eh_residente(livro, b.profissional)
                      for p in b.procedimentos if regras.mapear_convenio(p.convenio) == 'oci')
        self.assertGreater(raw_res, 0, 'o teste precisa de OCI em agenda de residente')

        views._oci_residentes(rel, livro)
        rel.blocos = [b for b in rel.blocos if not regras.eh_residente(livro, b.profissional)]

        movidas = [(b, p) for b in rel.blocos for p in b.procedimentos
                   if 'residente' in (p.motivo_calculo or '').lower()]
        self.assertEqual(len(movidas), raw_res, 'todo OCI de residente deve ser transferido (nada perdido)')
        # todas no bloco do Alessander e com honorário calculado (> 0)
        for b, p in movidas:
            self.assertIn('alessander', regras.normalizar(b.profissional))
            self.assertGreater(p.honorario or 0, 0)


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
        self.assertEqual(medplus.subclasse_preview(p('Facectomia', medplus.CLASSE_TAXA)),
                         medplus.SUBCLASSE_TAXA)


@unittest.skipUnless(_arquivos_presentes(PLANILHA, F22), 'amostra 22/06 ausente')
class CorrecaoPorMedicoTests(TestCase):
    """Correção memorizada por médico: vale só p/ o médico salvo (e os marcados)."""

    def test_isolada_por_medico(self):
        from repasses.services import correcoes
        from repasses import views
        correcoes.memorizar('Consulta em Oftalmologia', 'Particular', 999.0,
                            medico='Dr. Heric Sakamoto')
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(F22)
        views._filtrar_blocos(rel); views._separar_por_dia(rel)
        regras.processar(rel, livro)
        correcoes.aplicar(rel)

        def consultas_particular(frag):
            return [p.honorario for b in rel.blocos if frag in b.profissional.lower()
                    for p in b.procedimentos
                    if p.procedimento.strip().lower() == 'consulta em oftalmologia'
                    and (p.convenio or '').strip().lower() == 'particular']
        her = consultas_particular('heric')
        rob = consultas_particular('roberta')
        self.assertTrue(her and all(abs(h - 999.0) < 0.01 for h in her))   # Heric: todas 999
        self.assertTrue(rob and all(abs(h - 999.0) > 0.01 for h in rob))   # Roberta: nenhuma 999

    def test_casa_por_cadastro_mesmo_nome_diferente(self):
        # Bug pego na revisão: o modal lista o nome do CADASTRO (ex.: "Dra. Tharcila"),
        # que difere do nome do MedPlus ("Dra. Tharcila Breginski da Rocha"). Sem
        # resolver pelo cadastro, a correção do "outro médico" nunca casaria.
        from repasses.services import correcoes
        from repasses import views
        livro = regras.carregar_regras(PLANILHA)
        rel = medplus.ler_relatorio(F22)
        views._filtrar_blocos(rel); views._separar_por_dia(rel)
        regras.processar(rel, livro)
        b = next(x for x in rel.blocos if 'tharcila' in x.profissional.lower())
        canon = livro.medico_por_nome(b.profissional).nome
        self.assertNotEqual(canon, b.profissional)         # cadastro != MedPlus
        proc = next(p.procedimento for p in b.procedimentos if 'facoemulsifica' in p.procedimento.lower())
        correcoes.memorizar(proc, 'Sus Maringá', 777.0, medico=canon)   # como o modal salvaria

        rel2 = medplus.ler_relatorio(F22)
        views._filtrar_blocos(rel2); views._separar_por_dia(rel2)
        regras.processar(rel2, livro)
        correcoes.aplicar(rel2, livro)                     # resolve o médico pelo cadastro
        thar = [p.honorario for x in rel2.blocos if 'tharcila' in x.profissional.lower()
                for p in x.procedimentos if 'facoemulsifica' in p.procedimento.lower()
                and (p.convenio or '').strip().lower().startswith('sus')]
        self.assertTrue(thar and all(abs(h - 777.0) < 0.01 for h in thar))


class FellowSplitTests(TestCase):
    """Catarata com fellow: 60/40 calculado e EDITÁVEL (override do chefe e do fellow)."""

    def _rel_catarata(self, valor=1000.0):
        from types import SimpleNamespace
        p = medplus.Procedimento(None, '', 'Paciente X', 'Facoemulsificação', 'Particular', 1, valor, None)
        p.idx = 5
        p.status_calculo = regras.CATARATA
        b = medplus.BlocoMedico(profissional='Dr. Chefe')
        b.procedimentos = [p]
        return SimpleNamespace(blocos=[b]), p

    def _fellow_line(self, rel):
        return next(x for bl in rel.blocos for x in bl.procedimentos if getattr(x, 'sintetica', False))

    def test_calculado_60_40(self):
        from repasses import views
        rel, p = self._rel_catarata(1000.0)   # à vista 30% -> total 300; chefe 180, fellow 120
        views._resolver_cirurgias(rel, {'cat_modo_5': 'avista', 'cat_fellow_5': 'Dr. Fellow'})
        self.assertAlmostEqual(p.honorario, 180.0)
        self.assertAlmostEqual(self._fellow_line(rel).honorario, 120.0)

    def test_chefe_override_no_proprio_campo(self):
        from repasses import views
        rel, p = self._rel_catarata(10000.0)   # à vista 30% -> total 3000
        views._resolver_cirurgias(rel, {'cat_modo_5': 'avista', 'cat_fellow_5': 'Dr. Fellow',
                                        'cat_chefe_5': '2000'})
        self.assertAlmostEqual(p.honorario, 2000.0)                       # chefe sobrescrito
        self.assertAlmostEqual(self._fellow_line(rel).honorario, 1200.0)  # fellow segue 40%

    def test_fellow_override_na_propria_linha(self):
        from repasses import views
        rel, p = self._rel_catarata(10000.0)   # total 3000; chefe 1800, fellow 1200
        post = {'cat_modo_5': 'avista', 'cat_fellow_5': 'Dr. Fellow'}
        views._resolver_cirurgias(rel, post)
        views._indexar_sinteticas(rel)            # dá idx à linha do fellow
        fl = self._fellow_line(rel)
        post[f'hon_{fl.idx}'] = '900'             # edita o fellow NA LINHA DELE
        views._aplicar_edicoes_sinteticas(rel, post)
        self.assertAlmostEqual(fl.honorario, 900.0)
        self.assertAlmostEqual(p.honorario, 1800.0)  # chefe segue 60% (independente)

    def test_chefe_negativo_ignorado(self):
        from repasses import views
        rel, p = self._rel_catarata(10000.0)   # -500 inválido -> usa 60%
        views._resolver_cirurgias(rel, {'cat_modo_5': 'avista', 'cat_fellow_5': 'Dr. Fellow',
                                        'cat_chefe_5': '-500'})
        self.assertAlmostEqual(p.honorario, 1800.0)


class ReginaTests(SimpleTestCase):
    """Dra. Regina: valor por faixa + distinção indivíduo x equipe."""

    def test_valor_por_faixa(self):
        from repasses import views
        self.assertEqual(views._valor_regina(1), 1000.0)
        self.assertEqual(views._valor_regina(23), 1000.0)
        self.assertEqual(views._valor_regina(24), 1500.0)   # a partir de 24
        self.assertEqual(views._valor_regina(40), 1500.0)

    def test_individuo_vs_equipe(self):
        from repasses import views
        self.assertTrue(views._eh_regina_dinheiro('Dra. Regina'))
        self.assertFalse(views._eh_regina_dinheiro('Equipe Dra. Regina'))
        self.assertFalse(views._eh_regina_dinheiro('Redientes Dra. Regina'))
        self.assertFalse(views._eh_regina_dinheiro('Dra. Isabela Miwa Maeda'))


class CamposObrigatoriosTests(SimpleTestCase):
    """Campos a verificar (classe/forma de pagamento/fellow/anestesista) bloqueiam o export."""

    def _rel(self):
        from types import SimpleNamespace
        p = medplus.Procedimento(None, '', 'Pac', 'Facoemulsificação', 'Particular', 1, 1000, None)
        p.idx = 1
        p.classe = medplus.CLASSE_CIRURGIA      # tem_cirurgia -> exige anestesista
        p.status_calculo = regras.CATARATA      # exige forma de pagamento + fellow
        b = medplus.BlocoMedico(profissional='Dr. X'); b.procedimentos = [p]
        return SimpleNamespace(blocos=[b])

    def test_pendente_quando_nao_confirmado(self):
        from repasses import views
        campos = dict(views._campos_pendentes(self._rel(), {}))
        self.assertIn('anest_nome_0', campos)
        self.assertIn('cat_modo_1', campos)
        self.assertIn('cat_fellow_1', campos)

    def test_ok_quando_confirmado(self):
        from repasses import views
        post = {'anest_nome_0': '__sem__', 'cat_modo_1': 'avista', 'cat_fellow_1': '__sem__'}
        self.assertEqual(views._campos_pendentes(self._rel(), post), [])


class IndexSinteticasTests(SimpleTestCase):
    """idx ÚNICO p/ linhas sintéticas (preceptoria/fellow) — sem isso elas ficavam com
    idx=0 e colidiam com a 1a linha real (hon_0 duplicado corrompia o honorário no Salvar)."""

    def test_detecta_medico_fora_do_cadastro(self):
        from types import SimpleNamespace
        from repasses import views
        from repasses.services.regras import LivroRegras, Medico as MedicoR
        livro = LivroRegras(medicos=[MedicoR(nome='Dr. Heric Sakamoto', categoria='medico')])
        rel = SimpleNamespace(blocos=[medplus.BlocoMedico(profissional='Dr. Heric Sakamoto'),
                                      medplus.BlocoMedico(profissional='Dra. Fulana Nova Silva')])
        novos = views._medicos_desconhecidos(rel, livro)
        self.assertIn('Dra. Fulana Nova Silva', novos)     # fora do cadastro
        self.assertNotIn('Dr. Heric Sakamoto', novos)      # já cadastrado

    def test_idx_unico_sem_colisao(self):
        from types import SimpleNamespace
        from repasses import views

        def proc(nome, idx=0, sintetica=False):
            p = medplus.Procedimento(None, '', '', nome, '', 1, None, None)
            p.idx, p.sintetica = idx, sintetica
            return p

        b1 = medplus.BlocoMedico(profissional='Dr. A')
        b1.procedimentos = [proc('X', 0), proc('Y', 1), proc('Preceptoria', 0, sintetica=True)]
        b2 = medplus.BlocoMedico(profissional='Dr. B (fellow)')
        b2.procedimentos = [proc('Z', 2), proc('Participação fellow', 0, sintetica=True)]
        rel = SimpleNamespace(blocos=[b1, b2])

        views._indexar_sinteticas(rel)
        idxs = [p.idx for b in rel.blocos for p in b.procedimentos]
        self.assertEqual(len(idxs), len(set(idxs)), 'nenhum idx pode repetir (senão hon_ colide)')
        self.assertEqual([b1.procedimentos[0].idx, b1.procedimentos[1].idx], [0, 1])  # reais intactos
        # sintéticas receberam idx acima do maior real (2)
        self.assertGreater(b1.procedimentos[2].idx, 2)
        self.assertGreater(b2.procedimentos[1].idx, 2)


@unittest.skipUnless(_arquivos_presentes(F22), 'amostra 22/06 ausente')
class LoteHistoricoTests(TestCase):
    """Histórico: reabrir/editar/excluir um lote; fixado (pago) bloqueia."""

    TOKEN = 'a1b2c3d4e5f6a7b8c9d0e1f2a3b4c5d6.xls'

    def _cria_lote(self):
        from repasses.models import Lote, Repasse
        with open(F22, 'rb') as f:
            conteudo = f.read()
        lote = Lote.objects.create(token=self.TOKEN, arquivo_nome='medplus_22_06.xls',
                                   unidade='Maringá', upload_conteudo=conteudo, edicoes={})
        Repasse.objects.create(lote=lote, medico='Dr. Heric Sakamoto',
                               clinica='Maringá - Matriz', valor=100, status='gerado')
        return lote

    def test_reabrir_renderiza_revisao(self):
        from repasses.models import RepasseRascunho
        lote = self._cria_lote()
        r = self.client.post(f'/lotes/{lote.id}/reabrir/')
        self.assertEqual(r.status_code, 200)
        self.assertContains(r, 'Editando o lote')
        self.assertContains(r, self.TOKEN)            # form de revisão com o token do lote
        self.assertTrue(RepasseRascunho.objects.filter(token=self.TOKEN).exists())

    def test_pago_bloqueia_reabrir_e_excluir(self):
        from repasses.models import Lote
        lote = self._cria_lote()
        rp = lote.repasses.first(); rp.status = 'pago'; rp.save()
        self.assertTrue(lote.fixado)
        r = self.client.post(f'/lotes/{lote.id}/reabrir/', follow=True)
        self.assertContains(r, 'fixado')
        self.client.post(f'/lotes/{lote.id}/excluir/')
        self.assertTrue(Lote.objects.filter(id=lote.id).exists())   # não excluiu

    def test_excluir_remove_em_cascata(self):
        from repasses.models import Lote, Repasse
        lote = self._cria_lote()
        r = self.client.post(f'/lotes/{lote.id}/excluir/', follow=True)
        self.assertFalse(Lote.objects.filter(id=lote.id).exists())
        self.assertFalse(Repasse.objects.filter(lote_id=lote.id).exists())

    def test_cadastrar_medico_novo(self):
        from repasses.models import Medico
        lote = self._cria_lote()   # garante o upload disponível (Lote.upload_conteudo)
        nome = 'Dr. Freelancer Teste da Silva'
        # sem categoria escolhida -> NÃO cadastra (o sistema não assume)
        self.client.post('/importar/cadastrar-medicos/', {
            'token': lote.token, 'novo_count': '1',
            'novo_nome_0': nome, 'novo_categoria_0': ''})
        self.assertFalse(Medico.objects.filter(nome=nome).exists())
        # com categoria + papel -> cadastra
        self.client.post('/importar/cadastrar-medicos/', {
            'token': lote.token, 'novo_count': '1', 'novo_nome_0': nome,
            'novo_categoria_0': 'medico', 'novo_razao_0': 'FREELANCER LTDA',
            'novo_regra_0': 'R$ 100/consulta', 'novo_papeis_0': 'anestesista'})
        m = Medico.objects.filter(nome=nome).first()
        self.assertIsNotNone(m)
        self.assertEqual(m.categoria, 'medico')
        self.assertEqual(m.razao_social, 'FREELANCER LTDA')
        self.assertTrue(m.eh_anestesista)

    def test_baixar_tudo_zip(self):
        import zipfile
        from repasses.models import ArquivoSaida
        lote = self._cria_lote()
        ArquivoSaida.objects.create(lote=lote, grupo='Importação OMIE',
                                    nome='OMIE_Contas_a_Pagar.xlsx', conteudo=b'PK-fake-1')
        ArquivoSaida.objects.create(lote=lote, grupo='Repasses por médico',
                                    nome='Dr. Heric 22-06.pdf', conteudo=b'PDF-fake-2')
        r = self.client.get(f'/lotes/{lote.id}/baixar-zip/')
        self.assertEqual(r.status_code, 200)
        conteudo = b''.join(r.streaming_content)
        with zipfile.ZipFile(io.BytesIO(conteudo)) as zf:
            nomes = zf.namelist()
        self.assertEqual(len(nomes), 2)
        self.assertTrue(any('Contas_a_Pagar' in n for n in nomes))
        self.assertTrue(any(n.endswith('.pdf') for n in nomes))
