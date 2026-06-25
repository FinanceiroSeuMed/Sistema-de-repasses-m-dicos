# -*- coding: utf-8 -*-
"""Gera o relatório de DEMONSTRAÇÃO usado para validar o sistema (e o .exe).

Não é um arquivo da MedPlus de verdade: é um "copia e cola" curado dos casos de
repasse mais singulares já encontrados (catarata particular do Dr. Rodolpho com
fellow, OCI de residente -> Dr. Alessander, OCI da Equipe Dr. Keiti -> Dr. Keiti,
pacote R$1.000 do Keiti, cirurgias plásticas, taxa de sala, "a classificar",
procedimentos padrão, várias filiais/departamentos e PR2 x PR3).

Reproduz o layout REAL da MedPlus ("Procedimentos pela Agenda"): linha
"Profissional:" (col 0) + nome (col 31), cabeçalho com os rótulos nas colunas
6/37/67/107/122/142/167/177/207 e as linhas de dados nas mesmas colunas.

Requer xlwt (só para gerar; não é dependência do sistema):  pip install xlwt
Uso:  python exemplo/gerar_exemplo_demo.py [caminho_de_saida.xls]
Sem argumento, escreve "Exemplo - Procedimentos pela Agenda (casos de teste).xls"
ao lado deste script.
"""
import os
import sys
import xlwt

# Colunas reais do export da MedPlus (índices fixos, planilha "esparsa").
C_DATA, C_PAC, C_PROC, C_VAL, C_HON, C_CONV, C_QTD, C_CLIN, C_HORA = (
    6, 37, 67, 107, 122, 142, 167, 177, 207)
C_PROF_LABEL, C_PROF_NOME = 0, 31

DATA = '16/06/2026'
DATA2 = '17/06/2026'

# Filiais (o texto tem de bater com o de-para de Departamento da OMIE).
MATRIZ = 'Maringá - Matriz'
MANDAGUACU = 'Mandaguaçu'
PAICANDU = 'Paiçandu'
SARANDI = 'Sarandi'
AV_BRASIL = 'Maringá - Av Brasil'
MANDAGUARI = 'Mandaguari'
PR = 'Maringá - Filial PR2 e PR3'

# Cada agenda: (nome_do_profissional, [ (data, paciente, proc, valor, convenio, qtd, clinica, hora) ])
AGENDAS = [
    # --- Dr. Rodolpho: catarata PARTICULAR com fellow (carro-chefe) + cirurgia SUS,
    #     injeção, taxa de sala e consulta. Tudo na Matriz (Departamento 01).
    ('Dr. Rodolpho Takaishi Nanin Matsumoto', [
        (DATA, 'Mauro Andrade (catarata)', 'Facectomia com lente intra-ocular - AO', 6000.0, 'Particular', 1, MATRIZ, '08:00'),
        (DATA, 'Sonia Prado (catarata)', 'Facectomia com lente intra-ocular - Mono', 3500.0, 'Particular', 1, MATRIZ, '08:30'),
        (DATA, 'Jose Ribeiro (SUS)', 'FACOEMULSIFICACAO C/ IMPLANTE DE LIO', 771.60, 'Sus Maringá', 1, MATRIZ, '09:00'),
        (DATA, 'Marta Lima', 'Injeção Intravitrea com Avastin', 1800.0, 'Particular', 1, MATRIZ, '09:30'),
        (DATA, 'Mauro Andrade (catarata)', 'Facectomia com lente intra-ocular', 4765.0, 'Taxas De Sala', 1, MATRIZ, '08:00'),
        (DATA, 'Helena Souza', 'Consulta em Oftalmologia', 150.0, 'Particular', 1, MATRIZ, '10:00'),
    ]),
    # --- Residente: OCI feito por residente NÃO é dele -> vira repasse do Dr. Alessander.
    ('Dr. Daniel Magalhães', [
        (DATA, 'Paciente OCI 1', 'OCI AVALIAÇÃO INICIAL EM OFTALMOLOGIA - A PARTIR DE 9 ANOS', 160.0, 'OCI - SUS', 1, SARANDI, '08:00'),
        (DATA, 'Paciente OCI 2', 'OCI AVALIAÇÃO INICIAL EM OFTALMOLOGIA - A PARTIR DE 9 ANOS', 160.0, 'OCI - SUS', 1, SARANDI, '08:10'),
        (DATA, 'Paciente OCI 3', 'OCI AVALIAÇÃO INICIAL EM OFTALMOLOGIA - A PARTIR DE 9 ANOS', 160.0, 'OCI - SUS', 1, SARANDI, '08:20'),
        (DATA, 'Joana Reavaliação', 'CONSULTA EM OFTALMOLOGIA - GERAL', 13.37, 'Sus Maringá', 1, SARANDI, '08:30'),
    ]),
    # --- Equipe Dr. Keiti: os OCI lançados na agenda da equipe são repassados ao
    #     Dr. Keiti (NOVA regra). Clínica PR2/PR3 sem "PR3"/glaucoma no nome -> PR2.
    ('Equipe - Dr. Keiti Shirasu', [
        (DATA, 'Estrabismo 1', 'OCI AVALIAÇÃO DE ESTRABISMO', 200.0, 'OCI - SUS', 1, PR, '13:00'),
        (DATA, 'Estrabismo 2', 'OCI AVALIAÇÃO DE ESTRABISMO', 200.0, 'OCI - SUS', 1, PR, '13:10'),
        (DATA, 'Estrabismo 3', 'OCI AVALIAÇÃO DE ESTRABISMO', 200.0, 'OCI - SUS', 1, PR, '13:20'),
        (DATA, 'Estrabismo 4', 'OCI AVALIAÇÃO DE ESTRABISMO', 200.0, 'OCI - SUS', 1, PR, '13:30'),
        (DATA, 'Estrabismo 5', 'OCI AVALIAÇÃO DE ESTRABISMO', 200.0, 'OCI - SUS', 1, PR, '13:40'),
    ]),
    # --- Dr. Keiti: pacote de R$ 1.000 (consultas/exames do dia) + 30% da cirurgia.
    ('Dr. Keiti Fernando Shirasu', [
        (DATA, 'Carlos Eberle', 'Consulta em Oftalmologia', 150.0, 'Particular', 1, MANDAGUACU, '14:00'),
        (DATA, 'Vera Tavares', 'Consulta em Oftalmologia', 150.0, 'Particular', 1, MANDAGUACU, '14:20'),
        (DATA, 'Luis Prado', 'Consulta em Oftalmologia', 150.0, 'Particular', 1, MANDAGUACU, '14:40'),
        (DATA, 'Rita Campos', 'Capsulotomia YAG LASER - AO', 880.0, 'Particular', 1, MANDAGUACU, '15:00'),
    ]),
    # --- Dra. Tharcila: cirurgias plásticas/oculoplástica + um "a classificar" + taxa.
    ('Dra. Tharcila Breginski da Rocha', [
        (DATA, 'Ana Blefaro', 'BLEFAROPLASTIA SUPERIOR OU INFERIOR', 0.0, 'Cisamusep', 1, AV_BRASIL, '08:00'),
        (DATA, 'Pedro Tumor', 'TUMOR DE PALPEBRAS MONOCULAR - CIRURGIA', 0.0, 'Cisamusep', 1, AV_BRASIL, '08:40'),
        (DATA, 'Lucia Conjuntiva', 'EXERESE DE TUMOR DE CONJUNTIVA', 0.0, 'Sus Maringá', 1, AV_BRASIL, '09:20'),
        (DATA, 'Mario Pterigio', 'Pterígio com Autotransplante Conjuntival', 2500.0, 'Particular', 1, AV_BRASIL, '10:00'),
        (DATA, 'Ivo Catarata CISA', 'Cirurgia de Catarata - AO (Incluso: Consulta)', 0.0, 'Cisamusep', 1, AV_BRASIL, '10:20'),
        (DATA, 'Cleusa Cisto', 'Cisto Pequeno', 1450.0, 'Particular', 1, AV_BRASIL, '10:40'),
    ]),
    # --- Dr. Heric: procedimentos padrão (consultas/exames), Qtd>1, procedimento e
    #     um convênio "Copel" (deve ser tratado como Particular). Filial Mandaguari (06).
    ('Dr. Heric Sakamoto', [
        (DATA, 'Bianca Reis', 'Consulta em Oftalmologia', 150.0, 'Particular', 1, MANDAGUARI, '08:00'),
        (DATA, 'Otavio Nunes', 'Consulta em Oftalmologia', 140.0, 'Parcerias I', 1, MANDAGUARI, '08:15'),
        (DATA, 'Sueli Maia', 'Mapeamento de Retina - AO', 170.0, 'Particular', 1, MANDAGUARI, '08:30'),
        (DATA, 'Gilberto Sá', 'Tonometria - AO', 22.23, 'Amil', 2, MANDAGUARI, '08:45'),
        (DATA, 'Nair Lopes', 'CONSULTA EM OFTALMOLOGIA C/ EXAME', 51.71, 'Cisamusep', 1, MANDAGUARI, '09:00'),
        (DATA, 'Paulo Capsu', 'Capsulotomia YAG LASER - Mono', 480.0, 'Particular', 1, MANDAGUARI, '09:15'),
        (DATA, 'Edna Copel', 'Consulta em Oftalmologia', 150.0, 'Copel', 1, MANDAGUARI, '09:30'),
        (DATA, 'Nadia Sanepar', 'Consulta em Oftalmologia', 150.0, 'Sanepar', 1, MANDAGUARI, '09:45'),
        # "A classificar" COM valor (palpite não acha a classe, mas a regra tem preço):
        # cai em INDEFINIDA com honorário > 0 -> BLOQUEIA o export até classificar.
        (DATA, 'Rui Corpo Estranho', 'Retirada de Corpo Estranho da Córnea', 100.0, 'Particular', 1, MANDAGUARI, '10:00'),
    ]),
    # --- Filial Paiçandu (Departamento 03): consultas/exames padrão. Garante que
    #     o de-para de TODAS as 7 filiais (e o CNPJ do a receber) seja exercitado.
    ('Dra. Ana Paula Calil Guermandi', [
        (DATA, 'Olga Paiçandu', 'Consulta em Oftalmologia', 150.0, 'Particular', 1, PAICANDU, '08:00'),
        (DATA, 'Ivan Paiçandu', 'CONSULTA EM OFTALMOLOGIA C/ EXAME', 51.71, 'Cisamusep', 1, PAICANDU, '08:20'),
    ]),
    # --- Glaucoma/PR3: nome da agenda com "(PR3)" -> Departamento 07. PR3 (a do
    #     mesmo CNPJ sem PR3/glaucoma cai em 07. PR2).
    ('Dra. Licia Inazawa (PR3)', [
        (DATA2, 'Wagner Glaucoma', 'CASO NOVO - CONSULTA DE GLAUCOMA', 125.52, 'PG - SUS', 1, PR, '08:00'),
        (DATA2, 'Cida Pressão', 'Tonometria - AO', 22.23, 'Amil', 1, PR, '08:30'),
        (DATA2, 'Beto Glaucoma', 'Consulta em Oftalmologia', 150.0, 'Particular', 1, PR, '09:00'),
    ]),
]


def gerar(saida):
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Procedimentos pela Agenda')
    ws.write(0, C_PROF_NOME, 'SeuMed — Relatório de Demonstração')
    ws.write(2, C_PROF_NOME, 'Procedimentos pela Agenda')
    r = 4
    for nome, linhas in AGENDAS:
        ws.write(r, C_PROF_LABEL, 'Profissional:')
        ws.write(r, C_PROF_NOME, nome)
        r += 1
        # cabeçalho
        ws.write(r, C_DATA, 'Data Agend.'); ws.write(r, C_PAC, 'Nome Paciente')
        ws.write(r, C_PROC, 'Procedimento'); ws.write(r, C_VAL, 'Valor')
        ws.write(r, C_HON, 'Honorários'); ws.write(r, C_CONV, 'Convênio')
        ws.write(r, C_QTD, 'Qtd.'); ws.write(r, C_CLIN, 'Nome da Clínica')
        ws.write(r, C_HORA, 'Hora Agend.')
        r += 1
        for (data, pac, proc, val, conv, qtd, clin, hora) in linhas:
            ws.write(r, C_DATA, data); ws.write(r, C_PAC, pac)
            ws.write(r, C_PROC, proc); ws.write(r, C_VAL, float(val))
            ws.write(r, C_HON, 0.0); ws.write(r, C_CONV, conv)
            ws.write(r, C_QTD, int(qtd)); ws.write(r, C_CLIN, clin)
            ws.write(r, C_HORA, hora)
            r += 1
        r += 1  # linha em branco entre agendas
    wb.save(saida)
    print('Gerado:', saida, '| agendas:', len(AGENDAS),
          '| linhas:', sum(len(l) for _, l in AGENDAS))


if __name__ == '__main__':
    destino = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        'Exemplo - Procedimentos pela Agenda (casos de teste).xls')
    gerar(destino)
