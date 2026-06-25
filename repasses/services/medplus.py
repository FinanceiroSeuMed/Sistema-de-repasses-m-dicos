# -*- coding: utf-8 -*-
"""
Leitor (parser) do relatório "Procedimentos pela Agenda" exportado da MedPlus.

O arquivo NÃO é uma tabela limpa: é um relatório impresso, com o nome do médico,
uma linha de cabeçalho e as linhas de procedimentos espalhadas por colunas fixas.
O parser se orienta pelos RÓTULOS do cabeçalho (e não por posições fixas), para
resistir a mudanças de layout e suportar arquivos com vários médicos.
"""

from __future__ import annotations

import functools
import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd

# --- Classes de repasse -------------------------------------------------------

CLASSE_CIRURGIA = 'Cirurgias e Procedimentos'
CLASSE_EXAME = 'Exames e Consultas'
CLASSE_PRECEPTORIA = 'Preceptoria'
CLASSE_TAXA = 'Taxas de Sala'
CLASSE_INDEFINIDA = 'A classificar'

CLASSES = [CLASSE_CIRURGIA, CLASSE_EXAME, CLASSE_PRECEPTORIA, CLASSE_TAXA, CLASSE_INDEFINIDA]

# Subclasses só para a VISUALIZAÇÃO (preview do site): "Cirurgias e Procedimentos"
# se divide em Cirurgias vs Procedimentos (consultório: YAG/laser etc.). Na OMIE a
# pagar continua sendo uma classe só ("Cirurgias e Procedimentos").
SUBCLASSE_CIRURGIA = 'Cirurgias'
SUBCLASSE_PROCEDIMENTO = 'Procedimentos'
SUBCLASSE_EXAME = 'Exames e Consultas'
SUBCLASSE_PRECEPTORIA = 'Preceptorias'
SUBCLASSE_TAXA = 'Taxas de Sala'
SUBCLASSE_INDEFINIDA = 'A classificar'
SUBCLASSES = [SUBCLASSE_CIRURGIA, SUBCLASSE_PROCEDIMENTO, SUBCLASSE_EXAME,
              SUBCLASSE_PRECEPTORIA, SUBCLASSE_TAXA, SUBCLASSE_INDEFINIDA]

# Palavras-chave para um PALPITE inicial de classificação. É provisório:
# a classificação definitiva virá das regras (anexo 5) e poderá ser editada
# pelo usuário linha a linha.
_PALAVRAS_CLASSE = {
    CLASSE_CIRURGIA: [
        'facoemulsificacao', 'capsulotomia', 'yag', 'iridotomia', 'iridectomia',
        'calazio', 'pterigio', 'blefaroplastia', 'ptose', 'entropio', 'ectropio',
        'reconstrucao', 'tumor', 'exerese', 'injecao', 'intravitrea', 'avastin',
        'vitrectomia', 'transplante', 'cirurgia', 'sondagem', 'cantoplastia',
    ],
    CLASSE_EXAME: [
        'consulta', 'avaliacao', 'mapeamento', 'tonometria', 'gonioscopia',
        'oct', 'tomografia', 'retinografia', 'angiofluor', 'estereopsia',
        'campimetria', 'biometria', 'paquimetria', 'retina', 'exame', 'binocular',
        'fotocoagulacao', 'panfotocoagulacao', 'topografia', 'microscopia',
    ],
}

# Sinônimos aceitos para cada coluna do relatório (sem acento, minúsculo).
_ROTULOS = {
    'data': ('data agend', 'data'),
    'paciente': ('nome paciente', 'paciente'),
    'procedimento': ('procedimento',),
    'valor': ('valor',),
    'honorario': ('honorario',),
    'convenio': ('convenio',),
    'qtd': ('qtd', 'quantidade'),
    'clinica': ('nome da clinica', 'clinica', 'nome clinica', 'unidade', 'filial'),
    'hora': ('hora agend', 'horario', 'hora'),
}

_RE_DATA = re.compile(r'^\s*(\d{2})/(\d{2})/(\d{4})\s*$')


# --- Estruturas de dados ------------------------------------------------------

@dataclass
class Procedimento:
    data: date | None
    data_texto: str
    paciente: str
    procedimento: str
    convenio: str
    quantidade: int
    valor: float | None              # valor bruto (só existe no relatório completo)
    honorario_medplus: float | None  # honorário que veio da MedPlus (referência)
    classe: str = CLASSE_INDEFINIDA
    clinica: str = ''                # filial / "Nome da Clínica"
    hora: str = ''                   # "Hora Agend." — só para o preview (ajuda a achar o anestesista)
    # Preenchidos pelo motor de cálculo (regras.py)
    honorario: float | None = None   # honorário recalculado pelas regras
    status_calculo: str = ''         # calculado / nao_recebe / a_definir
    motivo_calculo: str = ''
    idx: int = 0                     # índice estável da linha (para edição na revisão)
    eh_catarata_part: bool = False   # catarata particular resolvida (mantém o seletor na tela)
    sintetica: bool = False          # linha criada na revisão (preceptoria / participação do fellow):
                                     # valor derivado/digitado à parte, renderizada como só-leitura
    editavel: bool = False           # linha sintética COM campo de honorário próprio (ex.: fellow)


@dataclass
class BlocoMedico:
    profissional: str
    procedimentos: list[Procedimento] = field(default_factory=list)

    @property
    def total_honorario_medplus(self) -> float:
        return round(sum((p.honorario_medplus or 0) for p in self.procedimentos), 2)

    @property
    def total_registros(self) -> int:
        # não conta componentes de cirurgia (anestesista/hospital) duplicados
        return sum(1 for p in self.procedimentos if p.status_calculo != 'componente')

    @property
    def resumo_classes(self) -> list[tuple[str, int]]:
        """Lista (classe, quantidade) na ordem padrão, só das classes presentes."""
        contagem = {}
        for p in self.procedimentos:
            contagem[p.classe] = contagem.get(p.classe, 0) + 1
        return [(c, contagem[c]) for c in CLASSES if c in contagem]

    @property
    def qtd_a_classificar(self) -> int:
        return sum(1 for p in self.procedimentos if p.classe == CLASSE_INDEFINIDA)

    @property
    def total_honorario(self) -> float:
        return round(sum((p.honorario or 0) for p in self.procedimentos
                         if p.status_calculo == 'calculado'), 2)

    @property
    def tem_cirurgia(self) -> bool:
        """Só cirurgias de fato (não procedimentos de consultório) ativam o anestesista."""
        return any(p.classe == CLASSE_CIRURGIA and eh_cirurgia(p.procedimento)
                   for p in self.procedimentos)

    @property
    def qtd_a_definir(self) -> int:
        return sum(1 for p in self.procedimentos if p.status_calculo == 'a_definir')

    @property
    def totais_por_classe(self) -> list[tuple[str, float]]:
        """(classe, total de honorários calculados) — para o preview."""
        tot = {}
        for p in self.procedimentos:
            if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                tot[p.classe] = tot.get(p.classe, 0) + p.honorario  # precisão cheia
        return [(c, round(tot[c], 2)) for c in CLASSES if c in tot]  # arredonda só no fim

    @property
    def linhas_agrupadas(self) -> list[tuple[str, list]]:
        """[(subclasse, [procedimentos])] na ordem do preview, só grupos não vazios.
        Divide 'Cirurgias e Procedimentos' em Cirurgias × Procedimentos."""
        grupos = {}
        for p in self.procedimentos:
            grupos.setdefault(subclasse_preview(p), []).append(p)
        return [(s, grupos[s]) for s in SUBCLASSES if s in grupos]

    @property
    def totais_por_subclasse(self) -> list[tuple[str, float]]:
        """(subclasse, total calculado) — badges do preview (4 tipos)."""
        tot = {}
        for p in self.procedimentos:
            if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                s = subclasse_preview(p)
                tot[s] = tot.get(s, 0) + p.honorario
        return [(s, round(tot[s], 2)) for s in SUBCLASSES if s in tot]

    # Preenchidos pelo orquestrador (regras.processar)
    lembrete: str = ''
    razao_social: str = ''
    data: date | None = None          # dia do bloco (após separar por dia)
    clinica: str = ''                 # filial do bloco (após separar por clínica)
    subunidade: str = ''              # 'PR2'/'PR3' (só p/ a clínica "PR2 e PR3"): glaucoma=PR3
    preceptoria_valor: float | None = None  # valor de preceptoria semanal sugerido
    participacao: bool = False        # bloco só de participação do fellow em catarata


@dataclass
class ResultadoImportacao:
    unidade: str
    blocos: list[BlocoMedico] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    anestesistas: list = field(default_factory=list)  # repasses de anestesista

    @property
    def total_registros(self) -> int:
        return sum(b.total_registros for b in self.blocos)


class ErroLeituraMedPlus(Exception):
    """Erro de leitura/validação do arquivo da MedPlus."""


# --- Funções auxiliares -------------------------------------------------------

def _norm(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ''
    texto = str(valor).strip().lower()
    return ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))


def _texto(valor) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ''
    return str(valor).strip()


def _para_numero(valor):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace('R$', '').strip()
    if not texto:
        return None
    # aceita "1.020,60" (pt-BR) e "1020.6" (ponto decimal)
    if ',' in texto and '.' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    try:
        return float(texto)
    except ValueError:
        return None


def _para_data(valor):
    if isinstance(valor, (datetime, pd.Timestamp)):
        d = valor.date() if hasattr(valor, 'date') else valor
        return d, d.strftime('%d/%m/%Y')
    texto = _texto(valor)
    m = _RE_DATA.match(texto)
    if m:
        dia, mes, ano = (int(g) for g in m.groups())
        try:
            d = date(ano, mes, dia)
            return d, d.strftime('%d/%m/%Y')
        except ValueError:
            return None, texto
    return None, texto


# Procedimentos da classe "Cirurgias e Procedimentos" que são CIRURGIAS de fato
# (precisam de anestesista). Os de consultório (YAG/laser/capsulotomia/iridotomia)
# NÃO entram no repasse do anestesista.
_CIRURGIAS_ANESTESIA = (
    'facoemulsificacao', 'facectomia', 'catarata', 'vitrectomia', 'blefaroplastia',
    'ptose', 'entropio', 'ectropio', 'pterigio', 'transplante', 'injecao', 'intravitrea',
    'trabeculectomia', 'reconstrucao', 'tumor', 'exerese', 'cantoplastia', 'sondagem',
    'recobrimento', 'enucleacao', 'evisceracao', 'calazio', 'cirurgia',
)


@functools.lru_cache(maxsize=2048)
def eh_cirurgia(procedimento: str) -> bool:
    """True se o procedimento é uma CIRURGIA (precisa de anestesista), distinguindo
    das de consultório (YAG/laser, capsulotomia/iridotomia a laser). Memoizado —
    é chamado por linha em tem_cirurgia/subclasse_preview e em regras.py."""
    n = _norm(procedimento)
    if any(k in n for k in ('yag', 'laser')):
        return False
    return any(k in n for k in _CIRURGIAS_ANESTESIA)


def subclasse_preview(p) -> str:
    """Subclasse só para o preview (Cirurgias × Procedimentos × Exames × Preceptorias)."""
    if p.classe == CLASSE_PRECEPTORIA:
        return SUBCLASSE_PRECEPTORIA
    if p.classe == CLASSE_TAXA:
        return SUBCLASSE_TAXA
    if p.classe == CLASSE_EXAME:
        return SUBCLASSE_EXAME
    if p.classe == CLASSE_CIRURGIA:
        return SUBCLASSE_CIRURGIA if eh_cirurgia(p.procedimento) else SUBCLASSE_PROCEDIMENTO
    return SUBCLASSE_INDEFINIDA


def classificar(procedimento: str, convenio: str = '') -> str:
    """Palpite inicial de classe (provisório).

    Usa o convênio (ex.: "Taxas de Sala" é lançado como convênio no MedPlus) e,
    em seguida, palavras-chave do nome do procedimento. A classe definitiva virá
    das regras (anexo 5) e poderá ser corrigida pelo usuário.
    """
    conv = _norm(convenio)
    if 'taxa' in conv:
        return CLASSE_TAXA
    n = _norm(procedimento)
    if not n:
        return CLASSE_INDEFINIDA
    for classe, palavras in _PALAVRAS_CLASSE.items():
        if any(p in n for p in palavras):
            return classe
    return CLASSE_INDEFINIDA


def _mapear_colunas(linha) -> dict | None:
    """Dada uma linha (Series), devolve {campo: indice} se for o cabeçalho."""
    indices = {}
    for j, celula in enumerate(linha):
        n = _norm(celula)
        if not n:
            continue
        for campo, sinonimos in _ROTULOS.items():
            if campo in indices:
                continue
            if any(n.startswith(s) for s in sinonimos):
                indices[campo] = j
    # é cabeçalho se tem ao menos procedimento + honorário + data
    if {'procedimento', 'honorario', 'data'} <= indices.keys():
        return indices
    return None


def _nome_profissional(linha) -> str:
    partes = []
    for celula in linha:
        n = _norm(celula)
        if not n or n.startswith('profissional'):
            continue
        partes.append(_texto(celula))
    return ' '.join(partes).strip()


# --- Função principal ---------------------------------------------------------

def ler_relatorio(arquivo, nome_arquivo: str = '') -> ResultadoImportacao:
    """
    Lê um relatório da MedPlus (caminho, bytes ou file-like) e devolve a
    estrutura com os médicos e seus procedimentos já com palpite de classe.
    """
    # Normaliza para bytes (caminho, bytes ou file-like) e detecta o formato REAL
    # pela assinatura — não confia só na extensão (um .xls pode ser OOXML).
    if hasattr(arquivo, 'read'):
        conteudo = arquivo.read()
    elif isinstance(arquivo, (bytes, bytearray)):
        conteudo = bytes(arquivo)
    else:
        with open(arquivo, 'rb') as _f:
            conteudo = _f.read()
    # OOXML (.xlsx) começa com 'PK\x03\x04'; senão tenta o .xls legado primeiro.
    engines = ('openpyxl', 'xlrd') if conteudo[:4] == b'PK\x03\x04' else ('xlrd', 'openpyxl')
    df, ultimo = None, None
    for engine in engines:
        try:
            df = pd.read_excel(io.BytesIO(conteudo), sheet_name=0, engine=engine, header=None)
            break
        except Exception as exc:   # tenta a outra engine antes de desistir
            ultimo = exc
    if df is None:
        raise ErroLeituraMedPlus(
            'Não consegui abrir o arquivo — ele não parece ser a planilha exportada da '
            'MedPlus (Excel .xls ou .xlsx). Reexporte pela MedPlus e tente novamente.'
        ) from ultimo

    resultado = ResultadoImportacao(unidade='')
    # unidade = primeiro texto não-vazio das primeiras linhas
    for i in range(min(5, len(df))):
        for celula in df.iloc[i]:
            t = _texto(celula)
            if t:
                resultado.unidade = t
                break
        if resultado.unidade:
            break

    colmap = None
    bloco_atual = None

    for i in range(len(df)):
        linha = df.iloc[i]
        textos = [_norm(c) for c in linha]
        junto = ' '.join(t for t in textos if t)

        if not junto:
            continue
        if junto.startswith('profissional') or any(t.startswith('profissional') for t in textos):
            nome = _nome_profissional(linha)
            if nome:
                bloco_atual = BlocoMedico(profissional=nome)
                resultado.blocos.append(bloco_atual)
            continue

        novo_cab = _mapear_colunas(linha)
        if novo_cab:
            colmap = novo_cab
            continue

        # linha de rodapé/total -> ignora
        if any(chave in junto for chave in ('gerado por', 'total de registros', 'pagina', 'página')):
            continue

        # candidata a linha de dados
        if colmap is None or bloco_atual is None:
            continue
        col_data = colmap.get('data')
        col_proc = colmap.get('procedimento')
        data_obj, data_txt = _para_data(linha.iloc[col_data]) if col_data is not None else (None, '')
        procedimento = _texto(linha.iloc[col_proc]) if col_proc is not None else ''
        if not procedimento or data_obj is None:
            continue  # sem data válida + procedimento não é linha de dado
        if _norm(procedimento).startswith('status'):
            continue  # linha de status do paciente (metadado do MedPlus), não é procedimento

        def _v(campo):
            idx = colmap.get(campo)
            return linha.iloc[idx] if idx is not None else None

        qtd = _para_numero(_v('qtd'))
        proc = Procedimento(
            data=data_obj,
            data_texto=data_txt,
            paciente=_texto(_v('paciente')),
            procedimento=procedimento,
            convenio=_texto(_v('convenio')),
            quantidade=int(qtd) if qtd else 1,
            valor=_para_numero(_v('valor')),
            honorario_medplus=_para_numero(_v('honorario')),
            clinica=_texto(_v('clinica')),
            hora=_texto(_v('hora')),
        )
        proc.classe = classificar(procedimento, proc.convenio)
        bloco_atual.procedimentos.append(proc)

    if not resultado.blocos or resultado.total_registros == 0:
        raise ErroLeituraMedPlus(
            'Não encontrei procedimentos no arquivo. Confirme que é o relatório '
            '"Procedimentos pela Agenda" exportado da MedPlus.'
        )

    return resultado
