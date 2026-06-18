# -*- coding: utf-8 -*-
"""
Geradores do repasse do médico: Excel (arquivo interno) e PDF (enviado ao médico).

- Um arquivo por médico.
- Mostra o honorário RECALCULADO (não o valor bruto do convênio).
- Nome do arquivo: "Dr. {nome} {DD-MM}" (data do atendimento).
- Lembrete de preceptoria (semanal/mensal) no topo, quando houver.
"""

from __future__ import annotations

import io
import re
import unicodedata

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

_INVALIDO = re.compile(r'[\\/:*?"<>|]+')


def moeda(valor) -> str:
    if valor is None:
        return 'A definir'
    texto = f'{float(valor):,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
    return f'R$ {texto}'


def _data_bloco(bloco):
    datas = [p.data for p in bloco.procedimentos if p.data]
    return max(datas) if datas else None


def nome_base(bloco) -> str:
    """Ex.: 'Dr. Heric Sakamoto 16-06' (sem caracteres inválidos em arquivo)."""
    nome = _INVALIDO.sub('', bloco.profissional).strip().rstrip('.').strip()
    data_ref = _data_bloco(bloco)
    dia = data_ref.strftime('%d-%m') if data_ref else 'sem-data'
    return f'{nome} {dia}'.strip()


def pagaveis(bloco):
    """Só as linhas com honorário a receber (> 0). Linhas R$ 0,00 e 'a definir'
    não entram no Excel arquivado nem no PDF enviado ao médico."""
    return [p for p in bloco.procedimentos
            if p.status_calculo == 'calculado' and (p.honorario or 0) > 0]


def _total(bloco) -> float:
    return round(sum(p.honorario for p in pagaveis(bloco)), 2)


# --- Excel (arquivo interno) --------------------------------------------------

def gerar_excel(bloco, unidade: str) -> bytes:
    data_ref = _data_bloco(bloco)
    wb = Workbook()
    ws = wb.active
    ws.title = 'Repasse'

    negrito = Font(bold=True)
    titulo = Font(bold=True, size=14)
    cabec_fill = PatternFill('solid', fgColor='0B5FA5')
    cabec_font = Font(bold=True, color='FFFFFF')

    ws['A1'] = unidade or 'Repasse Médico'
    ws['A1'].font = titulo
    ws['A2'] = 'Demonstrativo de Repasse Médico'
    ws['A2'].font = negrito
    ws['A3'] = f'Profissional: {bloco.profissional}'
    if getattr(bloco, 'razao_social', ''):
        ws['A4'] = f'Razão Social: {bloco.razao_social}'
    if data_ref:
        ws['A5'] = f'Data do atendimento: {data_ref.strftime("%d/%m/%Y")}'
    linha = 6
    if getattr(bloco, 'lembrete', ''):
        ws.cell(linha, 1, f'⚠ {bloco.lembrete}').font = Font(italic=True, color='946200')
        linha += 1
    linha += 1

    colunas = ['Data', 'Paciente', 'Procedimento', 'Convênio', 'Qtd.', 'Honorário']
    for c, titulo_col in enumerate(colunas, start=1):
        cel = ws.cell(linha, c, titulo_col)
        cel.fill = cabec_fill
        cel.font = cabec_font
    linha += 1

    for p in pagaveis(bloco):
        ws.cell(linha, 1, p.data_texto)
        ws.cell(linha, 2, p.paciente)
        ws.cell(linha, 3, p.procedimento)
        ws.cell(linha, 4, p.convenio)
        ws.cell(linha, 5, p.quantidade)
        cel_h = ws.cell(linha, 6, p.honorario)
        cel_h.number_format = 'R$ #,##0.00'
        linha += 1

    ws.cell(linha, 5, 'Total').font = negrito
    cel_total = ws.cell(linha, 6, _total(bloco))
    cel_total.number_format = 'R$ #,##0.00'
    cel_total.font = negrito

    larguras = [12, 26, 44, 16, 6, 14]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --- PDF (enviado ao médico) --------------------------------------------------

def gerar_pdf(bloco, unidade: str) -> bytes:
    data_ref = _data_bloco(bloco)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=18 * mm, bottomMargin=16 * mm,
                            leftMargin=16 * mm, rightMargin=16 * mm,
                            title=f'Repasse {bloco.profissional}')
    estilos = getSampleStyleSheet()
    st_titulo = ParagraphStyle('t', parent=estilos['Title'], fontSize=15, spaceAfter=2,
                               textColor=colors.HexColor('#0B5FA5'))
    st_sub = ParagraphStyle('s', parent=estilos['Normal'], fontSize=9, textColor=colors.HexColor('#52606d'))
    st_info = ParagraphStyle('i', parent=estilos['Normal'], fontSize=10, spaceAfter=2)
    st_lembrete = ParagraphStyle('l', parent=estilos['Normal'], fontSize=10,
                                 textColor=colors.HexColor('#946200'), backColor=colors.HexColor('#FFF7E6'),
                                 borderPadding=6, spaceBefore=6, spaceAfter=6)
    st_cel = ParagraphStyle('c', parent=estilos['Normal'], fontSize=8.5, leading=11)

    elems = [
        Paragraph(unidade or 'Repasse Médico', st_titulo),
        Paragraph('Demonstrativo de Repasse Médico', st_sub),
        Spacer(1, 8),
        Paragraph(f'<b>Profissional:</b> {bloco.profissional}', st_info),
    ]
    if getattr(bloco, 'razao_social', ''):
        elems.append(Paragraph(f'<b>Razão Social:</b> {bloco.razao_social}', st_info))
    if data_ref:
        elems.append(Paragraph(f'<b>Data do atendimento:</b> {data_ref.strftime("%d/%m/%Y")}', st_info))
    if getattr(bloco, 'lembrete', ''):
        elems.append(Paragraph(f'📌 {bloco.lembrete}', st_lembrete))
    elems.append(Spacer(1, 8))

    dados = [['Data', 'Procedimento', 'Convênio', 'Qtd.', 'Honorário']]
    for p in pagaveis(bloco):
        dados.append([
            p.data_texto,
            Paragraph(p.procedimento, st_cel),
            p.convenio,
            str(p.quantidade),
            moeda(p.honorario),
        ])
    dados.append(['', '', '', 'Total', moeda(_total(bloco))])

    tabela = Table(dados, colWidths=[22 * mm, 78 * mm, 30 * mm, 12 * mm, 28 * mm], repeatRows=1)
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0B5FA5')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('ALIGN', (3, 0), (4, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#F7F9FC')]),
        ('LINEBELOW', (0, 0), (-1, -1), 0.4, colors.HexColor('#E2E8F0')),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#EEF2F7')),
        ('FONTNAME', (3, -1), (4, -1), 'Helvetica-Bold'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    elems.append(tabela)
    elems.append(Spacer(1, 10))
    elems.append(Paragraph(
        f'Total de {len(bloco.procedimentos)} procedimento(s). '
        'Documento para conferência do profissional.', st_sub))

    doc.build(elems)
    return buf.getvalue()
