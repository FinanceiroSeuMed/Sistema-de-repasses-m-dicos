import io
import re
from collections import Counter
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from django.conf import settings
from django.contrib import messages
from django.db.models import Count, Q
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from .forms import ImportarMedPlusForm
from .models import (AjusteMensal, ArquivoSaida, ClasseMemorizada, CorrecaoMemorizada,
                     DiaSemRepasse, Lote, Medico, RegraRepasse, Repasse, RepasseRascunho)
from .services import correcoes, medplus, omie, regras, repasse


def home(request):
    """Tela inicial / painel do sistema de repasses médicos."""
    contexto = {
        'total_medicos': Medico.objects.count(),
        'total_medicos_ativos': Medico.objects.filter(ativo=True).count(),
        'edicao_em_andamento': _rascunho_em_andamento() is not None,
    }
    return render(request, 'repasses/home.html', contexto)


def _rascunho_em_andamento():
    """O rascunho do último import ainda em edição (com o arquivo disponível), ou None.
    As edições só são largadas ao importar outro arquivo ou após exportar."""
    for r in RepasseRascunho.objects.order_by('-atualizado_em'):
        if _caminho_upload(r.token) is not None:
            return r
    return None


def continuar_edicao(request):
    """Volta para a revisão do último import em edição — as edições ficam em cache
    (rascunho) enquanto se navega pelas outras abas. (Diretoria 2026-06-24.)"""
    r = _rascunho_em_andamento()
    if r is None:
        messages.info(request, 'Não há repasse em edição. Importe um arquivo para começar.')
        return redirect('repasses:importar')
    try:
        resultado, aviso, dados = _preparar_revisao(r.token)
    except medplus.ErroLeituraMedPlus as exc:
        messages.error(request, f'Não foi possível reabrir a edição: {exc}')
        return redirect('repasses:importar')
    return render(request, 'repasses/revisao.html',
                  _ctx_revisao(resultado, r.token, aviso, edicoes=dados))


def medicos(request):
    """Cadastro central de médicos, agrupado por categoria."""
    todos = Medico.objects.all()
    grupos = []
    for codigo, rotulo in Medico.CATEGORIA_CHOICES:
        qs = [m for m in todos if m.categoria == codigo]
        if qs:
            grupos.append((rotulo, qs))
    contexto = {
        'grupos': grupos,
        'total': len(todos),
        'total_ativos': sum(1 for m in todos if m.ativo),
    }
    return render(request, 'repasses/medicos.html', contexto)


# --- Upload temporário (aguardando revisão) -----------------------------------

_TOKEN_RE = re.compile(r'^[0-9a-f]{32}\.(xls|xlsx)$')


def _salvar_upload(arquivo) -> str:
    ext = '.xlsx' if (arquivo.name or '').lower().endswith('.xlsx') else '.xls'
    token = uuid4().hex + ext
    destino = Path(settings.UPLOADS_DIR)
    destino.mkdir(parents=True, exist_ok=True)
    (destino / token).write_bytes(arquivo.read())
    return token


def _token_base(token: str) -> str:
    """Token do UPLOAD a partir do token de um lote: os lotes são por DIA
    ('<upload>~AAAA-MM-DD'), mas o arquivo importado/rascunho é um só."""
    return (token or '').split('~', 1)[0]


def _caminho_upload(token: str):
    token = _token_base(token)
    if not token or not _TOKEN_RE.match(token):
        return None
    base = Path(settings.UPLOADS_DIR).resolve()
    caminho = (base / token).resolve()
    if base not in caminho.parents:
        return None
    if caminho.is_file():
        return caminho
    # Pasta uploads/ limpa? Recupera o .xls de origem guardado no banco — qualquer
    # lote-dia deste upload serve (todos guardam o mesmo .xls).
    lote = (Lote.objects.filter(token__startswith=token)
            .exclude(upload_conteudo=None).first())
    if lote and lote.upload_conteudo:
        base.mkdir(parents=True, exist_ok=True)
        caminho.write_bytes(bytes(lote.upload_conteudo))
        return caminho
    return None


# --- Rascunho: memória das edições do repasse em andamento ---------------------

# Campos da revisão que ficam salvos (o resto — token, csrf, memorizar — não).
_PREFIXOS_RASCUNHO = ('hon_', 'classe_', 'cat_modo_', 'cat_fellow_',
                      'cat_chefe_', 'equipe_destino_', 'oci_integracao_',
                      'anest_nome_', 'anest_horas_', 'anest_cnpj_', 'preceptoria_', 'memo_medicos_')


def _salvar_rascunho(token, post):
    if not token:
        return
    # Se o POST não traz NENHUM campo do formulário de revisão (ex.: veio da tela
    # Visualizar, que é só-leitura), não sobrescreve — senão apagaria as edições.
    if not any(chave.startswith(_PREFIXOS_RASCUNHO) for chave in post.keys()):
        return
    dados = {}
    for chave in post.keys():
        if chave.startswith(_PREFIXOS_RASCUNHO):
            valor = (post.get(chave) or '').strip()
            if valor:
                dados[chave] = valor
    RepasseRascunho.objects.update_or_create(token=token, defaults={'dados': dados})


def _carregar_rascunho(token):
    r = RepasseRascunho.objects.filter(token=token).first()
    return dict(r.dados) if r else {}


def _anotar_selecoes(resultado, dados):
    """Marca em cada bloco/procedimento o que foi escolhido (para a tela não
    'esquecer' os selects de catarata, anestesista e preceptoria)."""
    for i, bloco in enumerate(resultado.blocos):
        bloco.sel_anest_nome = dados.get(f'anest_nome_{i}', '')
        bloco.sel_anest_horas = dados.get(f'anest_horas_{i}', '')
        bloco.sel_anest_cnpj = dados.get(f'anest_cnpj_{i}', '')
        bloco.sel_preceptoria = dados.get(f'preceptoria_{i}', '')
        bloco.sel_equipe_destino = dados.get(f'equipe_destino_{i}', '')
        bloco.sel_oci_integracao = dados.get(f'oci_integracao_{i}', '')
        for p in bloco.procedimentos:
            p.sel_cat_modo = dados.get(f'cat_modo_{p.idx}', '')
            p.sel_cat_fellow = dados.get(f'cat_fellow_{p.idx}', '')
            p.sel_cat_chefe = dados.get(f'cat_chefe_{p.idx}', '')


def _preparar_revisao(token):
    """Monta o resultado para a tela: lê o arquivo, reaplica as edições salvas e
    resolve anestesistas (blocos próprios), catarata particular e preceptoria — para
    que o que o usuário Salva já apareça no preview, igual ao que será exportado."""
    caminho = _caminho_upload(token)
    if caminho is None:
        return None, None, {}
    resultado, aviso = _ler_e_processar(caminho, token)
    resultado.log_edicoes = []          # auditoria: ajustes manuais de honorário
    dados = _carregar_rascunho(token)
    _aplicar_edicoes(resultado, dados)
    # Agenda "Equipe Dr. Keiti": roteia para o Dr. Keiti ou a Dra. Thalia (escolha da
    # diretoria). Roda ANTES dos demais resolvers (antes de criar blocos sintéticos),
    # para o índice do seletor casar com a posição do bloco. (Diretoria 2026-06-27.)
    _resolver_equipe(resultado, dados)
    # OCI de residente: registra a escolha (integrar no Dr. Alessander ou não). O move/
    # descarte efetivo é na exportação, para não mexer no índice dos campos.
    _resolver_oci_residentes(resultado, dados)
    _resolver_anestesistas(resultado, dados)
    # Resolve a catarata particular JÁ no preview: quando o usuário declara à
    # vista/parcelado, o valor entra no total e o repasse do fellow (40%) aparece.
    _resolver_cirurgias(resultado, dados)
    # Preceptoria semanal também é aplicada no preview: ao Salvar (não só com Enter
    # no campo), a linha de preceptoria já aparece no demonstrativo. (Diretoria 2026-06-23.)
    _resolver_preceptoria(resultado, dados)
    # idx único p/ as linhas sintéticas criadas acima (preceptoria/fellow) — evita
    # colisão de hon_/classe_ com as linhas reais no formulário.
    _indexar_sinteticas(resultado)
    # Honorário editado na PRÓPRIA linha do fellow (idx só existe após o passo acima).
    _aplicar_edicoes_sinteticas(resultado, dados)
    _anotar_selecoes(resultado, dados)
    return resultado, aviso, dados


def _aplicar_edicoes_sinteticas(resultado, post):
    """Aplica o honorário editado nas linhas sintéticas EDITÁVEIS (participação do
    fellow), que ganham idx próprio depois de criadas e têm campo na sua linha."""
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            if getattr(p, 'editavel', False):
                raw = (post.get(f'hon_{p.idx}') or '').strip()
                v = _num(raw)
                if v is not None and v >= 0:
                    p.honorario = v
                    p.sel_hon = raw     # mostra o valor digitado (preto); vazio -> placeholder cinza
                else:
                    p.sel_hon = ''


_RE_SUFIXO_CADASTRO = re.compile(r'\s*\([^)]*\)\s*$')
_RE_TITULO_MEDICO = re.compile(r'dra?\b')
# Médicos/agendas deixados de fora por ora (entram em funcionalidade futura)
_EXCLUIR_MEDICOS = ('crivari', 'guilherme', 'maria marta')

# Agenda "Equipe Dr. Keiti": o destino (Dr. Keiti OU Dra. Thalia) é escolhido pela
# diretoria na revisão (até existir uma agenda "Equipe Dra. Thalia"). (2026-06-27.)
_NOME_KEITI = 'Dr. Keiti Fernando Shirasu'
_NOME_EQUIPE_KEITI = 'Equipe Dr. Keiti'


def _eh_equipe_keiti(n):
    """A agenda é a 'Equipe Dr. Keiti'? `n` já vem normalizado (minúsculo, s/ acento)."""
    return 'equipe' in n and 'keiti' in n


_SEM_PRECEPTOR = '__sem_preceptor__'


def _equipe_destinos():
    """Preceptores que podem receber a agenda 'Equipe Dr. Keiti' — vêm do CADASTRO
    (eh_preceptor=True), então cadastrar/remover um preceptor reflete na lista na hora.
    (Diretoria 2026-07-01.)"""
    return list(Medico.objects.filter(eh_preceptor=True)
                .order_by('nome').values_list('nome', flat=True))


def _filtrar_blocos(resultado):
    """Remove agendas que não são de médico (sem Dr./Dra.) e os deixados de fora."""
    novos = []
    for bloco in resultado.blocos:
        n = regras.normalizar(bloco.profissional)
        if _eh_equipe_keiti(n):
            # Mantém a agenda (apesar de não começar com "Dr."); o destino é decidido
            # na revisão por _resolver_equipe. Marca para o resto do fluxo reconhecer.
            bloco.equipe_keiti = True
            bloco.profissional = _NOME_EQUIPE_KEITI
            novos.append(bloco)
            continue
        if not _RE_TITULO_MEDICO.match(n):
            continue  # ex.: "Agenda Glaucoma", "Agenda Externa"
        if any(x in n for x in _EXCLUIR_MEDICOS):
            continue  # Crivari, Guilherme, Maria Marta — por ora
        novos.append(bloco)
    resultado.blocos = novos


# Status especial: taxa de sala com valor em aberto. Aparece na revisão com campo
# de valor vazio; só entra no repasse/a pagar se o usuário informar um valor devido.
STATUS_TAXA_SALA = 'taxa_sala'


def _marcar_taxas_sala(resultado):
    """Taxas de sala ficam VISÍVEIS na revisão com valor em aberto (diretoria
    2026-06-23): o usuário informa o valor só se for devido ao médico; em branco,
    são desconsideradas no repasse. (Antes eram removidas da revisão.)"""
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            # Taxa de sala é identificada pelo CONVÊNIO (igual ao medplus.classificar) —
            # assim, mesmo que a memória por procedimento tenha mexido na classe, a
            # linha continua reconhecida como taxa. (Diretoria 2026-06-26.)
            if p.classe == medplus.CLASSE_TAXA or 'taxa' in regras.normalizar(p.convenio):
                # Respeita um valor já vindo de correção memorizada (correcoes.aplicar
                # roda antes): só deixa "em aberto" quem ainda não tem valor.
                if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                    continue
                p.status_calculo = STATUS_TAXA_SALA
                p.honorario = None
                p.motivo_calculo = ('Taxa de sala — informe o valor só se for devido '
                                    'ao médico; em branco, fica fora do repasse.')


def _linha_vale(p):
    """Linhas que NÃO entram em nenhum lugar (nem preview): R$0, componentes e
    'Não faturável'. As taxas de sala AGORA permanecem (valor opcional na revisão)."""
    if p.status_calculo in ('nao_recebe', 'componente'):
        return False
    if 'nao faturavel' in regras.normalizar(p.procedimento):
        return False
    return True


def _limpar_linhas(resultado):
    for bloco in resultado.blocos:
        bloco.procedimentos = [p for p in bloco.procedimentos if _linha_vale(p)]
    resultado.blocos = [b for b in resultado.blocos if b.procedimentos]


def _nome_base_medico(nome):
    """Remove o sufixo de unidade entre parênteses: 'Dr. Carlos (PR3)' -> 'Dr. Carlos'.
    As variantes (Geral/PR2/PR3) são a MESMA pessoa — a clínica vem da coluna própria."""
    return _RE_SUFIXO_CADASTRO.sub('', nome or '').strip()


def _ler_e_processar(caminho, nome=''):
    resultado = medplus.ler_relatorio(str(caminho), nome)
    _filtrar_blocos(resultado)
    # NÃO unificamos filiais (GERAL/PR2/PR3): cada filial gera seu próprio repasse;
    # o agrupamento por razão social acontece só na OMIE a pagar.
    _separar_por_dia(resultado)
    livro = regras.carregar_livro_padrao()
    aviso = None
    if livro is None:
        aviso = ('A planilha de regras não foi encontrada — os honorários ficaram '
                 '"a definir". Confira REGRAS_REPASSE_PATH.')
    else:
        regras.processar(resultado, livro)
        # OCI feito por residente -> o repasse é do Dr. Alessander (residente não recebe)
        _oci_residentes(resultado, livro)
        # Correções memorizadas: reaplicam ajustes manuais salvos em meses anteriores
        # (livro resolve o médico pelo cadastro p/ a correção por-médico casar)
        correcoes.aplicar(resultado, livro)
        # Classe memorizada por procedimento (vale p/ todos os médicos) — reaplica a
        # classificação que a diretoria já definiu antes. (Diretoria 2026-06-26.)
        correcoes.aplicar_classes(resultado)
        # Residentes não recebem -> não aparecem no preview nem na exportação, EXCETO os
        # que têm OCI a decidir (integrar no Dr. Alessander?) — ficam para a revisão.
        resultado.blocos = [b for b in resultado.blocos
                            if not regras.eh_residente(livro, b.profissional)
                            or getattr(b, 'oci_residente', False)]
        _aplicar_keiti(resultado)
        # Equipe Dr. Keiti: os lançamentos são SEMPRE consultas/exames (diretoria
        # 2026-06-27) — nunca cirurgia — no a pagar/receber da OMIE. Força após a
        # memória de classes para não ser sobrescrito.
        for bloco in resultado.blocos:
            if getattr(bloco, 'equipe_keiti', False):
                for p in bloco.procedimentos:
                    p.classe = medplus.CLASSE_EXAME
        _marcar_preceptoria(resultado, livro)
        _marcar_taxas_sala(resultado)
    _limpar_linhas(resultado)
    _indexar(resultado)
    # Guarda a classe SUGERIDA (após regras + memória) p/ detectar quando o usuário
    # muda a classe e, então, memorizar a nova classificação.
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            p.classe_sugerida = p.classe
    resultado.medicos_novos = _medicos_desconhecidos(resultado, livro)
    return resultado, aviso


def _medicos_desconhecidos(resultado, livro):
    """Nomes de médicos do import que NÃO estão no cadastro — precisam ser
    classificados pelo usuário (o sistema nunca assume a categoria)."""
    if livro is None:
        return []
    vistos, novos = set(), []
    for bloco in resultado.blocos:
        if getattr(bloco, 'participacao', False):     # bloco sintético (fellow) — ignora
            continue
        if getattr(bloco, 'equipe_keiti', False):     # "Equipe Dr. Keiti" — destino é escolhido
            continue
        if livro.medico_por_nome(bloco.profissional) is not None:
            continue
        nome = _nome_base_medico(bloco.profissional)
        chave = regras.normalizar(nome)
        if chave and chave not in vistos:
            vistos.add(chave)
            novos.append(nome)
    return novos


def _eh_pr3(nome) -> bool:
    """A agenda é de GLAUCOMA / PR3? (decide o Departamento na clínica 'PR2 e PR3'.)
    Regra da diretoria: nome da agenda com 'PR3' ou 'glaucoma' -> PR3; senão PR2."""
    n = omie._norm_clinica(nome)        # minúsculas, sem acento (mantém '(pr3)')
    return 'pr3' in n or 'glaucoma' in n


def _separar_por_dia(resultado):
    """Um bloco por (MÉDICO, DIA, CLÍNICA[, PR2/PR3]).

    O sufixo de unidade no nome ('(Geral)', '(PR3)') em geral NÃO diferencia: é a
    mesma pessoa. Mas na clínica 'Maringá - Filial PR2 e PR3' a SUBUNIDADE importa
    para o Departamento OMIE: agenda de glaucoma/PR3 -> PR3, o resto -> PR2; então
    aí PR2 e PR3 viram blocos separados."""
    from datetime import date as _date
    grupos = {}
    ordem = []
    for bloco in resultado.blocos:
        nome = _nome_base_medico(bloco.profissional)
        sub_origem = 'PR3' if _eh_pr3(bloco.profissional) else 'PR2'
        for p in bloco.procedimentos:
            sub = sub_origem if omie.clinica_pr2_pr3(p.clinica) else ''
            chave = (nome, p.data, p.clinica, sub)
            nb = grupos.get(chave)
            if nb is None:
                nb = medplus.BlocoMedico(profissional=nome)
                nb.data = p.data
                nb.clinica = p.clinica
                nb.subunidade = sub
                nb.equipe_keiti = getattr(bloco, 'equipe_keiti', False)  # preserva a marca
                grupos[chave] = nb
                ordem.append(chave)
            nb.procedimentos.append(p)

    def _chave(k):
        nome, d, c, sub = k
        return (d or _date.max, c or '', sub, nome)

    resultado.blocos = [grupos[k] for k in sorted(ordem, key=_chave)]


# Médico que recebe o repasse do OCI feito por residentes (casa por trecho do nome).
_OCI_RESPONSAVEL = 'alessander'


def _oci_residentes(resultado, livro):
    """OCI feito por residente: o residente NÃO recebe. A diretoria decide na revisão se
    INTEGRA o OCI no repasse do Dr. Alessander (Sim) ou não (Não = não paga). Aqui só
    MARCAMOS: mantém no bloco do residente APENAS as linhas de OCI (recalculadas como se o
    Dr. Alessander as tivesse feito) e sinaliza oci_residente=True. O move/descarte
    acontece na exportação (_aplicar_oci_residentes). Rodar ANTES da filtragem de
    residentes. (Diretoria 2026-07-01 — antes era automático.)"""
    responsavel = next((m for m in livro.medicos
                        if _OCI_RESPONSAVEL in regras.normalizar(m.nome)), None)
    if responsavel is None:
        return
    for bloco in resultado.blocos:
        if not regras.eh_residente(livro, bloco.profissional):
            continue
        ocis = []
        for p in bloco.procedimentos:
            if regras.mapear_convenio(p.convenio) == 'oci':
                r = regras.calcular(livro, p.procedimento, p.convenio, p.valor,
                                    responsavel.nome, medico_obj=responsavel,
                                    quantidade=p.quantidade)
                p.honorario = r.honorario
                p.status_calculo = r.status
                p.motivo_calculo = (f'OCI de residente ({bloco.profissional}) — a integrar '
                                    f'no {responsavel.nome} (decidir na revisão).')
                ocis.append(p)
        if ocis:   # só mantém o bloco do residente se tiver OCI a decidir
            bloco.procedimentos = ocis
            bloco.oci_residente = True
            bloco.oci_responsavel = responsavel.nome


def _aplicar_oci_residentes(resultado):
    """Na EXPORTAÇÃO: integra os OCI de residente no Dr. Alessander (oci_integra='sim')
    ou descarta (qualquer outro — não paga, residente não recebe). Depois remove os blocos
    de residente (só existiam para a decisão). Feito só ao exportar para não mexer no
    índice dos campos da revisão."""
    from collections import defaultdict
    mover, responsavel_nome = defaultdict(list), ''
    for bloco in resultado.blocos:
        if getattr(bloco, 'oci_residente', False) and getattr(bloco, 'oci_integra', '') == 'sim':
            responsavel_nome = getattr(bloco, 'oci_responsavel', '') or responsavel_nome
            for ln in bloco.procedimentos:
                # movida pelo SISTEMA para a agenda do responsável — se vier sem valor
                # bruto, não é incoerência do arquivo (não avisa). (Diretoria 2026-07-02.)
                ln.sem_bruto_sistema = True
            mover[(bloco.data, bloco.clinica or '')].extend(bloco.procedimentos)
    resultado.blocos = [b for b in resultado.blocos if not getattr(b, 'oci_residente', False)]
    if not mover:
        return
    m = Medico.objects.filter(nome=responsavel_nome).first()
    alvo = regras.normalizar(responsavel_nome)
    for (data, clinica), linhas in mover.items():
        bloco = next((b for b in resultado.blocos
                      if regras.normalizar(b.profissional) == alvo
                      and b.data == data and (b.clinica or '') == clinica), None)
        if bloco is None:
            bloco = medplus.BlocoMedico(profissional=responsavel_nome,
                                        razao_social=(m.razao_social if m else '') or '')
            bloco.data = data
            bloco.clinica = clinica
            bloco.cnpj = m.cnpj if m else ''
            bloco.medico_cadastro = responsavel_nome
            resultado.blocos.append(bloco)
        bloco.procedimentos.extend(linhas)


def _aplicar_keiti(resultado):
    """Dr. Keiti: R$ 1.000 (consultas/exames do dia) + 30% do valor das cirurgias.

    OCI (inclusive os vindos da agenda "Equipe Dr. Keiti") NÃO entram no pacote de
    R$ 1.000 — cada OCI mantém o seu próprio valor de repasse (regra do convênio)."""
    for bloco in resultado.blocos:
        # A agenda "Equipe Dr. Keiti" tem destino próprio (Keiti/Thalia, em
        # _resolver_equipe) e usa o valor por linha — não o pacote de R$ 1.000.
        if getattr(bloco, 'equipe_keiti', False):
            continue
        if 'keiti' not in regras.normalizar(bloco.profissional):
            continue
        tem_exame = False
        novas = []
        for p in bloco.procedimentos:
            if regras.mapear_convenio(p.convenio) == 'oci':
                # OCI da agenda "Equipe Dr. Keiti" -> repasse do Dr. Keiti, valor próprio.
                if not (p.motivo_calculo or '').lower().startswith('oci'):
                    p.motivo_calculo = (f'OCI repassado ao Dr. Keiti — {p.motivo_calculo}'
                                        if p.motivo_calculo else 'OCI repassado ao Dr. Keiti.')
                novas.append(p)
            elif p.classe == medplus.CLASSE_CIRURGIA and p.valor:
                p.honorario = round(0.30 * p.valor, 2)
                p.status_calculo = 'calculado'
                p.motivo_calculo = 'Dr. Keiti: 30% do valor da cirurgia.'
                novas.append(p)
            elif p.classe in (medplus.CLASSE_EXAME, medplus.CLASSE_INDEFINIDA):
                tem_exame = True  # absorvido no pacote de R$ 1.000
            else:
                novas.append(p)
        if tem_exame:
            ref = bloco.procedimentos[0] if bloco.procedimentos else None
            pacote = medplus.Procedimento(
                data=getattr(ref, 'data', None), data_texto=getattr(ref, 'data_texto', ''),
                paciente='', procedimento='Consultas e Exames (pacote Dr. Keiti)',
                convenio='', quantidade=1, valor=None, honorario_medplus=None,
                classe=medplus.CLASSE_EXAME)
            pacote.honorario = 1000.0
            pacote.status_calculo = 'calculado'
            pacote.motivo_calculo = 'Dr. Keiti: R$ 1.000 por consultas/exames do dia.'
            pacote.sem_bruto_sistema = True   # criado pelo sistema — sem bruto por natureza
            novas.append(pacote)
        bloco.procedimentos = novas


def _num_money(texto):
    import re as _re
    m = _re.search(r'r\$\s*([\d.]+,\d{2}|\d[\d.]*)', (texto or '').lower())
    return _num(m.group(1)) if m else None


def _marcar_preceptoria(resultado, livro):
    """Sugere o valor de preceptoria semanal (campo editável na revisão)."""
    for bloco in resultado.blocos:
        m = livro.medico_por_nome(bloco.profissional)
        if m and 'preceptor' in m.categoria.lower() and 'semana' in (m.obs or '').lower():
            bloco.preceptoria_valor = _num_money(m.obs)


def _indexar(resultado):
    """Atribui um índice estável a cada procedimento (para edição na revisão)."""
    i = 0
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            p.idx = i
            i += 1


def _indexar_sinteticas(resultado):
    """Dá um idx ÚNICO às linhas sintéticas (preceptoria / participação do fellow)
    criadas na revisão — sem mexer no idx das linhas reais (preserva o casamento do
    rascunho). Sem isso elas ficariam com idx=0 (default) e colidiriam com a 1a linha
    real (hon_0 duplicado no formulário, corrompendo o honorário no Salvar)."""
    reais = [p.idx for b in resultado.blocos for p in b.procedimentos if not p.sintetica]
    nxt = (max(reais) + 1) if reais else 0
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            if p.sintetica:
                mae = getattr(p, 'idx_mae', None)
                if mae is not None:
                    # ESTÁVEL: deriva da linha-mãe (ex.: participação do assistente ->
                    # 'f<idx da catarata>'). O sequencial mudava quando outra catarata
                    # virava "sem fellow", e o honorário editado do assistente migrava
                    # de linha. (Auditoria 2026-07-02.)
                    p.idx = f'f{mae}'
                else:
                    p.idx = nxt
                    nxt += 1


def _num(texto):
    t = (texto or '').strip().replace('R$', '').replace(' ', '')
    if not t:
        return None
    if ',' in t and '.' in t:
        t = t.replace('.', '').replace(',', '.')
    elif ',' in t:
        t = t.replace(',', '.')
    try:
        return round(float(t), 2)
    except ValueError:
        return None


def _aplicar_edicoes(resultado, post):
    """Aplica as edições feitas na tela de revisão (honorário e classe).

    O honorário calculado é mantido em precisão cheia (sem arredondar) para que a
    soma final feche centavo a centavo. A tela mostra o valor com 2 casas; se o
    usuário NÃO mexeu no campo, o que volta no POST é esse valor arredondado —
    então só sobrescrevemos quando ele difere do arredondado do original, senão
    perderíamos as casas extras a cada exportação.
    """
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            if p.status_calculo == regras.CATARATA:
                continue  # catarata é resolvida em _resolver_cirurgias
            classe = (post.get(f'classe_{p.idx}') or '').strip()
            if classe:
                p.classe = classe
            valor = _num(post.get(f'hon_{p.idx}'))
            if valor is None:
                continue
            original = p.honorario if p.honorario is not None else None
            if original is not None and round(float(original), 2) == round(float(valor), 2):
                continue  # inalterado — preserva a precisão cheia do cálculo
            antes = 'R$ %.2f' % float(original) if original is not None else '—'
            resultado.log_edicoes.append(
                f'Honorário: {bloco.profissional} — "{p.procedimento[:40]}" '
                f'({p.convenio or "s/ convênio"}) {antes} → R$ %.2f' % valor)
            p.honorario = valor
            p.status_calculo = 'calculado' if valor > 0 else 'nao_recebe'
            p.motivo_calculo = 'Editado manualmente na revisão.'


def _resolver_cirurgias(resultado, post):
    """Resolve a catarata particular: à vista/parcelado, split 60/40 com fellow,
    e cria as linhas do fellow (que aparecem no repasse dele e do cirurgião)."""
    from collections import defaultdict
    extras = defaultdict(list)
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            if p.status_calculo != regras.CATARATA or p.valor is None:
                continue
            modo = (post.get(f'cat_modo_{p.idx}') or '').strip()
            fellow = (post.get(f'cat_fellow_{p.idx}') or '').strip()
            if fellow in (_VERIFICAR, _SEM):
                fellow = ''   # "a verificar"/"sem fellow" -> cirurgião 100%
            if modo not in ('avista', 'parcelado'):
                continue  # forma de pagamento não confirmada — permanece pendente
            taxa = regras.CATARATA_AVISTA if modo == 'avista' else regras.CATARATA_PARCELADO
            total = taxa * p.valor
            # O CIRURGIÃO edita o valor no seu próprio campo (cat_chefe_), com o valor
            # pré-calculado como placeholder cinza — uma única linha, sem linha extra de
            # resultado. Vale com OU sem assistente. (Diretoria 2026-06-26.)
            chefe_ovr = _num(post.get(f'cat_chefe_{p.idx}'))
            manual_chefe = chefe_ovr is not None and chefe_ovr >= 0
            if fellow:
                # Split 60/40: cirurgião 60% (campo dele) + assistente 40% (campo na
                # PRÓPRIA linha de participação, aplicado em _aplicar_edicoes_sinteticas).
                chefe_calc = round(total * (1 - regras.FELLOW_PERCENTUAL), 2)
                fellow_calc = round(round(total, 2) - chefe_calc, 2)
                p.chefe_calc = chefe_calc                          # placeholder do cat_chefe
                p.honorario = chefe_ovr if manual_chefe else chefe_calc
                p.sufixo_export = ' - Cirurgião 60%'               # adendo no repasse exportado
                p.motivo_calculo = (f'Catarata particular ({modo}) — Cirurgião 60%; Assistente {fellow} 40%'
                                    + (' (cirurgião ajustado).' if manual_chefe else '.'))
                linha = medplus.Procedimento(
                    data=p.data, data_texto=p.data_texto, paciente=p.paciente,
                    procedimento=p.procedimento,
                    convenio=p.convenio, quantidade=p.quantidade, valor=p.valor,
                    honorario_medplus=None, classe=medplus.CLASSE_CIRURGIA, hora=p.hora)
                linha.honorario = fellow_calc       # 40% padrão (editável na própria linha)
                linha.fellow_calc = fellow_calc     # placeholder cinza do campo do assistente
                linha.status_calculo = 'calculado'
                linha.sintetica = True
                linha.editavel = True               # tem campo de honorário próprio
                linha.idx_mae = p.idx               # idx ESTÁVEL: deriva da catarata-mãe
                linha.sufixo_export = ' - Assistente 40%'
                linha.motivo_calculo = f'Assistente 40% da catarata de {bloco.profissional}.'
                extras[(fellow, p.data, p.clinica)].append(linha)
            else:
                p.chefe_calc = round(total, 2)                     # placeholder (100%)
                p.honorario = chefe_ovr if manual_chefe else round(total, 2)
                p.motivo_calculo = (f'Catarata particular ({modo}) — Cirurgião 100% (sem assistente)'
                                    + (' (ajustado).' if manual_chefe else '.'))
            p.status_calculo = 'calculado'
            p.eh_catarata_part = True   # mantém o seletor à vista/parcelado na tela

    for (fellow, data, clinica), linhas in extras.items():
        bloco = next((b for b in resultado.blocos
                      if regras.normalizar(b.profissional) == regras.normalizar(fellow)
                      and b.data == data and (b.clinica or '') == (clinica or '')), None)
        if bloco is None:
            m = (Medico.objects.filter(nome=fellow).first()
                 or Medico.objects.filter(nome__iexact=fellow).first())
            bloco = medplus.BlocoMedico(profissional=fellow,
                                        razao_social=(m.razao_social if m else ''))
            bloco.data = data
            bloco.clinica = clinica
            # CNPJ do cadastro: sem ele o a pagar OMIE avisava "sem CNPJ" para o
            # assistente (ex.: Dr. Carlos Eduardo) mesmo com o cadastro correto.
            bloco.cnpj = m.cnpj if m else ''
            bloco.medico_cadastro = m.nome if m else fellow
            bloco.participacao = True   # bloco só de participação em catarata (sem caixa de anestesista)
            resultado.blocos.append(bloco)
        bloco.procedimentos.extend(linhas)


def _resolver_equipe(resultado, post):
    """A agenda 'Equipe Dr. Keiti' vai para o PRECEPTOR responsável (lista do cadastro,
    eh_preceptor) escolhido na revisão. 'Sem preceptor' = lançamento só para os residentes
    verem — não gera repasse (removido na hora de exportar). Sem escolha, fica pendente.
    Renomeia o bloco para o preceptor (nome + razão + CNPJ do cadastro). (Diretoria 2026-07-01.)"""
    validos = set(_equipe_destinos())
    for i, bloco in enumerate(resultado.blocos):
        if not getattr(bloco, 'equipe_keiti', False):
            continue
        destino = (post.get(f'equipe_destino_{i}') or '').strip()
        if destino == _SEM_PRECEPTOR:
            bloco.equipe_sem_preceptor = True
            bloco.equipe_resolvido = destino
            continue
        bloco.equipe_sem_preceptor = False
        if destino not in validos:
            continue   # não escolhido -> permanece pendente
        m = Medico.objects.filter(nome=destino).first()
        bloco.profissional = m.nome if m else destino
        bloco.razao_social = (m.razao_social if m else '') or bloco.razao_social
        bloco.cnpj = (m.cnpj if m else '') or getattr(bloco, 'cnpj', '')
        bloco.medico_cadastro = m.nome if m else destino
        bloco.equipe_resolvido = destino


def _resolver_oci_residentes(resultado, post):
    """Registra a escolha (integrar o OCI do residente no Dr. Alessander ou não). O move/
    descarte efetivo é em _aplicar_oci_residentes, na exportação."""
    for i, bloco in enumerate(resultado.blocos):
        if getattr(bloco, 'oci_residente', False):
            bloco.oci_integra = (post.get(f'oci_integracao_{i}') or '').strip()


def _resolver_preceptoria(resultado, post):
    """Adiciona a linha de preceptoria semanal informada pelo usuário na revisão."""
    for i, bloco in enumerate(list(resultado.blocos)):
        valor = _num(post.get(f'preceptoria_{i}'))
        if not valor or valor <= 0:
            continue
        linha = medplus.Procedimento(
            data=bloco.data, data_texto=(bloco.data.strftime('%d/%m/%Y') if bloco.data else ''),
            paciente='', procedimento='Preceptoria (semanal)', convenio='',
            quantidade=1, valor=None, honorario_medplus=None,
            classe=medplus.CLASSE_PRECEPTORIA)
        linha.honorario = round(valor, 2)
        linha.status_calculo = 'calculado'
        linha.sintetica = True   # linha derivada — só-leitura, idx único (não colide com hon_)
        linha.motivo_calculo = 'Repasse de preceptoria semanal.'
        bloco.procedimentos.append(linha)


# Sentinelas dos campos obrigatórios (devem ser confirmados antes de exportar):
_VERIFICAR = '__verificar__'   # pré-seleção "a verificar" (não confirmado)
_SEM = '__sem__'               # confirmado "sem" (sem anestesista / sem fellow)


def _eh_regina_dinheiro(nome):
    """A 'Dra. Regina' (indivíduo, paga em DINHEIRO, fora da OMIE) — mas NÃO a
    EQUIPE/Residentes da Dra. Regina, que recebem repasse normal. (Diretoria 2026-06-25.)"""
    n = regras.normalizar(nome)
    return 'regina' in n and not any(w in n for w in ('equipe', 'residente', 'rediente', 'grupo', 'time'))


def _eh_residente_regina(nome):
    """'Residente(s) Dra. Regina' — recebe repasse na OMIE, mas não tem cadastro/CNPJ
    próprio; por isso a diretoria digita o CNPJ na hora. (Diretoria 2026-07-01.)"""
    n = regras.normalizar(nome)
    return 'regina' in n and any(w in n for w in ('residente', 'rediente'))


def _valor_regina(n_cirurgias):
    """Valor em dinheiro da Dra. Regina por nº de cirurgias no dia: R$ 1.500 a partir
    de 24, senão R$ 1.000. (Diretoria 2026-06-25.)"""
    return 1500.0 if n_cirurgias >= 24 else 1000.0


def _resolver_anestesistas(resultado, post):
    """Para cada cirurgião com cirurgia, registra o anestesista escolhido na revisão
    (valor fixo do dia + horas extras) — gera linha no a pagar e repasse. A Dra. Regina
    é exceção: paga em dinheiro, só conta as cirurgias e vira um LEMBRETE (não OMIE)."""
    from collections import defaultdict
    regina = defaultdict(lambda: {'n': 0, 'clinicas': set()})   # dia -> contagem de cirurgias
    for i, bloco in enumerate(list(resultado.blocos)):
        if not bloco.tem_cirurgia:
            continue
        nome = (post.get(f'anest_nome_{i}') or '').strip()
        if not nome or nome in (_VERIFICAR, _SEM):
            continue  # não escolhido / sem anestesista (obrigatoriedade checada na exportação)
        # Só as CIRURGIAS de fato (não YAG/laser de consultório).
        cirurgias = [p for p in bloco.procedimentos
                     if p.classe == medplus.CLASSE_CIRURGIA and medplus.eh_cirurgia(p.procedimento)]
        datas = [p.data for p in cirurgias if p.data]
        dia = max(datas) if datas else bloco.data
        if _eh_regina_dinheiro(nome):
            regina[dia]['n'] += len(cirurgias)
            regina[dia]['clinicas'].add(getattr(bloco, 'clinica', '') or '')
            continue
        horas = _num(post.get(f'anest_horas_{i}')) or 0
        valor = regras.valor_anestesista(nome, horas)
        m = Medico.objects.filter(nome=nome).first()
        # 'Residente Dra. Regina' não tem CNPJ no cadastro: a diretoria digita na hora
        # (campo anest_cnpj_{i}) e ele vira o Fornecedor no a pagar. (Diretoria 2026-07-01.)
        cnpj_manual = (post.get(f'anest_cnpj_{i}') or '').strip()
        cnpj = cnpj_manual or (m.cnpj if m else '')
        resultado.anestesistas.append({
            'indice': i,
            'anestesista': nome,
            'razao_social': m.razao_social if m else '',
            'cnpj': cnpj,
            'cirurgiao': bloco.profissional,
            'clinica': getattr(bloco, 'clinica', '') or '',
            'subunidade': getattr(bloco, 'subunidade', ''),
            'data': dia,
            'horas': int(horas) if horas else 0,
            'valor': valor,
            'cirurgias': cirurgias,
        })
    # Lembretes da Dra. Regina (dinheiro, fora da OMIE) — um por dia.
    resultado.lembretes_regina = [
        {'dia': dia.strftime('%d/%m/%Y') if dia else 'sem data', 'n': info['n'],
         'valor': _valor_regina(info['n'])}
        for dia, info in sorted(regina.items(), key=lambda kv: str(kv[0])) if info['n']]


def _campos_pendentes(resultado, post):
    """Campos OBRIGATÓRIOS ainda não confirmados ('a verificar') — bloqueiam a
    exportação. Devolve [(nome_do_campo, descrição)] na ordem da tela:
    classificação (A classificar), forma de pagamento + fellow (catarata particular)
    e anestesista (cirurgias). (Diretoria 2026-06-25.)"""
    pend = []
    destinos_equipe = None   # consulta o cadastro só se houver agenda de equipe
    for i, bloco in enumerate(resultado.blocos):
        if getattr(bloco, 'equipe_keiti', False):
            if destinos_equipe is None:
                destinos_equipe = set(_equipe_destinos()) | {_SEM_PRECEPTOR}
            if (post.get(f'equipe_destino_{i}') or '').strip() not in destinos_equipe:
                pend.append((f'equipe_destino_{i}', 'Agenda "Equipe Dr. Keiti": escolha o preceptor '
                             'responsável (ou "Sem preceptor").'))
        if getattr(bloco, 'oci_residente', False):
            if (post.get(f'oci_integracao_{i}') or '').strip() not in ('sim', 'nao'):
                pend.append((f'oci_integracao_{i}', f'OCI na agenda de {bloco.profissional}: integrar '
                             'no Dr. Alessander? (Sim/Não)'))
        if bloco.tem_cirurgia and not getattr(bloco, 'participacao', False):
            nome_anest = (post.get(f'anest_nome_{i}') or '').strip()
            if nome_anest in ('', _VERIFICAR):
                pend.append((f'anest_nome_{i}', f'Anestesista de {bloco.profissional} — confirme (ou "sem anestesista").'))
            elif _eh_residente_regina(nome_anest) and not (post.get(f'anest_cnpj_{i}') or '').strip():
                pend.append((f'anest_cnpj_{i}', f'CNPJ da Residente Dra. Regina ({bloco.profissional}) — obrigatório para o a pagar.'))
        for p in bloco.procedimentos:
            if p.status_calculo == regras.CATARATA or getattr(p, 'eh_catarata_part', False):
                if (post.get(f'cat_modo_{p.idx}') or '').strip() not in ('avista', 'parcelado'):
                    pend.append((f'cat_modo_{p.idx}', f'Forma de pagamento da catarata ({bloco.profissional}).'))
                if (post.get(f'cat_fellow_{p.idx}') or '').strip() in ('', _VERIFICAR):
                    pend.append((f'cat_fellow_{p.idx}', f'Fellow na catarata ({bloco.profissional}) — confirme (ou "sem fellow").'))
            elif (p.classe == medplus.CLASSE_INDEFINIDA and (p.honorario or 0) > 0
                  and not getattr(p, 'sintetica', False)):
                pend.append((f'classe_{p.idx}', f'Classifique "{p.procedimento[:34]}".'))
    return pend


def _resumir_pendencias(itens):
    contagem = Counter(itens)
    return [f'{q}× {msg}' if q > 1 else msg for msg, q in contagem.items()]


def _pendencias_revisao(resultado):
    itens = []
    for b in resultado.blocos:
        tem_pagavel = any(p.status_calculo == 'calculado' and (p.honorario or 0) > 0
                          for p in b.procedimentos)
        if tem_pagavel and not b.razao_social:
            itens.append(f'{b.profissional}: sem Razão Social — confira na OMIE.')
        for p in b.procedimentos:
            if p.status_calculo == 'a_definir':
                itens.append(f'{b.profissional}: "{p.procedimento[:40]}" a definir.')
            elif p.status_calculo == regras.CATARATA:
                itens.append(f'{b.profissional}: catarata particular — informe à vista/parcelado e fellow.')
    return _resumir_pendencias(itens)


def _salvar_saidas(arquivos):
    """arquivos: lista de (grupo, nome_arquivo, conteudo). Grava e devolve (pasta, downloads)."""
    pasta = f'{datetime.now():%Y%m%d-%H%M%S}-{uuid4().hex[:6]}'
    destino = Path(settings.SAIDAS_DIR) / pasta
    destino.mkdir(parents=True, exist_ok=True)
    downloads = []
    for grupo, nome_arquivo, conteudo in arquivos:
        (destino / nome_arquivo).write_bytes(conteudo)
        downloads.append({'grupo': grupo, 'arquivo': nome_arquivo})
    return pasta, downloads


# --- Fluxo: importar -> revisar -> exportar -----------------------------------

def importar(request):
    """Passo 1: upload do relatório da MedPlus."""
    erro = None
    if request.method == 'POST':
        form = ImportarMedPlusForm(request.POST, request.FILES)
        if form.is_valid():
            arquivo = form.cleaned_data['arquivo']
            token = _salvar_upload(arquivo)
            try:
                resultado, aviso, dados = _preparar_revisao(token)
            except medplus.ErroLeituraMedPlus as exc:
                # Arquivo inválido NÃO destrói a edição em andamento do repasse anterior
                # (o rascunho antigo só é zerado após a leitura dar certo). (Auditoria
                # 2026-07-02.)
                erro = str(exc)
            else:
                # Novo repasse importado com sucesso -> zera a memória de edições do anterior.
                RepasseRascunho.objects.all().delete()
                RepasseRascunho.objects.create(token=token, arquivo_nome=(arquivo.name or ''), dados={})
                return render(request, 'repasses/revisao.html',
                              _ctx_revisao(resultado, token, aviso, edicoes=dados))
    else:
        form = ImportarMedPlusForm()
    return render(request, 'repasses/importar.html', {'form': form, 'erro': erro})


def salvar(request):
    """Salva as edições do repasse (rascunho) sem gerar arquivos — para a pessoa
    ir ajustando aos poucos sem perder nada."""
    if request.method != 'POST':
        raise Http404()
    token = request.POST.get('token', '')
    if _caminho_upload(token) is None:
        messages.error(request, 'Arquivo da importação não encontrado — refaça o upload.')
        return redirect('repasses:importar')
    _salvar_rascunho(token, request.POST)
    resultado, aviso, dados = _preparar_revisao(token)
    # Prévia dos arquivos (por médico + OMIE) + possíveis erros da OMIE, para a pessoa
    # conferir cobertura/erros logo ao salvar, sem precisar exportar. (Diretoria 2026-06-30.)
    previa = _previa_arquivos(resultado)
    previa_avisos = _resumir_pendencias(
        omie.gerar_contas_pagar(resultado, settings.OMIE_PAGAR_TEMPLATE).pendencias
        + omie.gerar_contas_receber(resultado, settings.OMIE_RECEBER_TEMPLATE).pendencias)
    info = ['Alterações salvas. Veja abaixo os arquivos que serão gerados (confira os '
            'médicos e os avisos). Pode continuar editando — ficam guardadas até importar outro.']
    return render(request, 'repasses/revisao.html',
                  _ctx_revisao(resultado, token, aviso, info=info, edicoes=dados,
                               previa=previa, previa_avisos=previa_avisos))


def cadastrar_medicos(request):
    """Cadastra os médicos novos (sem cadastro) classificados pelo usuário na revisão.

    O sistema NUNCA assume a categoria: só cadastra quem teve uma categoria escolhida.
    Campos editáveis (razão social, regra) para casos extraordinários. Depois recarrega
    a revisão — já reprocessada com os médicos cadastrados."""
    if request.method != 'POST':
        raise Http404()
    token = request.POST.get('token', '')
    if _caminho_upload(token) is None:
        messages.error(request, 'Arquivo da importação não encontrado — refaça o upload.')
        return redirect('repasses:importar')
    validos = dict(Medico.CATEGORIA_CHOICES)
    criados, sem_classe = 0, 0
    for i in range(int(request.POST.get('novo_count') or 0)):
        nome = (request.POST.get(f'novo_nome_{i}') or '').strip()
        categoria = (request.POST.get(f'novo_categoria_{i}') or '').strip()
        if not nome:
            continue
        if categoria not in validos:        # sem classificação -> NÃO cadastra (não assume)
            sem_classe += 1
            continue
        if Medico.objects.filter(nome__iexact=nome).exists():
            continue
        papeis = request.POST.getlist(f'novo_papeis_{i}')
        Medico.objects.create(
            nome=nome, categoria=categoria,
            razao_social=(request.POST.get(f'novo_razao_{i}') or '').strip(),
            nome_fantasia=(request.POST.get(f'novo_fantasia_{i}') or '').strip(),
            cnpj=(request.POST.get(f'novo_cnpj_{i}') or '').strip(),
            chave_pix=(request.POST.get(f'novo_pix_{i}') or '').strip(),
            regra_obs=(request.POST.get(f'novo_regra_{i}') or '').strip(),
            eh_fellow=(categoria == Medico.CATEGORIA_FELLOW) or ('fellow' in papeis),
            eh_preceptor=(categoria == Medico.CATEGORIA_PRECEPTOR) or ('preceptor' in papeis),
            eh_anestesista=(categoria == Medico.CATEGORIA_ANESTESISTA) or ('anestesista' in papeis))
        criados += 1
    if criados:
        messages.success(request, f'{criados} médico(s) cadastrado(s) — repasse reprocessado com eles.')
    if sem_classe:
        messages.error(request, f'{sem_classe} médico(s) sem categoria escolhida não foram cadastrados '
                       '— escolha a classificação (o sistema não assume).')
    resultado, aviso, dados = _preparar_revisao(token)
    return render(request, 'repasses/revisao.html', _ctx_revisao(resultado, token, aviso, edicoes=dados))


def revisar(request):
    """Volta para a tela de revisão (edição) sem salvar nem gerar nada —
    usado pelo botão 'Voltar à revisão' da tela de visualização."""
    if request.method != 'POST':
        raise Http404()
    token = request.POST.get('token', '')
    if _caminho_upload(token) is None:
        messages.error(request, 'Arquivo da importação não encontrado — refaça o upload.')
        return redirect('repasses:importar')
    resultado, aviso, dados = _preparar_revisao(token)
    return render(request, 'repasses/revisao.html',
                  _ctx_revisao(resultado, token, aviso, edicoes=dados))


def _resumo_visualizacao(resultado):
    """Contagens e totais para a tela de visualização (consulta rápida)."""
    contagem = Counter()
    total_pagar = 0.0
    medicos = set()
    for b in resultado.blocos:
        pagavel = False
        for p in b.procedimentos:
            if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                contagem[medplus.subclasse_preview(p)] += 1
                total_pagar += p.honorario
                pagavel = True
        if pagavel:
            medicos.add(b.profissional)
    total_pagar += sum(a.get('valor') or 0 for a in resultado.anestesistas)
    total_receber = sum(float(p.valor) for b in resultado.blocos for p in b.procedimentos
                        if p.valor is not None)
    return {
        'n_medicos': len(medicos),
        'cirurgias': contagem.get(medplus.SUBCLASSE_CIRURGIA, 0),
        'procedimentos': contagem.get(medplus.SUBCLASSE_PROCEDIMENTO, 0),
        'exames': contagem.get(medplus.SUBCLASSE_EXAME, 0),
        'preceptorias': contagem.get(medplus.SUBCLASSE_PRECEPTORIA, 0),
        'n_anestesistas': len(resultado.anestesistas),
        'total_pagar': round(total_pagar, 2),
        'total_receber': round(total_receber, 2),
    }


def visualizar(request):
    """Ação 'só visualizar': relatório estruturado, só-leitura, SEM gerar arquivo.
    Salva as edições atuais (vem da revisão com o formulário completo)."""
    if request.method != 'POST':
        raise Http404()
    token = request.POST.get('token', '')
    if _caminho_upload(token) is None:
        messages.error(request, 'Arquivo da importação não encontrado — refaça o upload.')
        return redirect('repasses:importar')
    _salvar_rascunho(token, request.POST)
    resultado, aviso, dados = _preparar_revisao(token)
    return render(request, 'repasses/visualizar.html', {
        'resultado': resultado,
        'token': token,
        'aviso_regras': aviso,
        'resumo': _resumo_visualizacao(resultado),
        'pendencias': _pendencias_revisao(resultado),
    })


def _ctx_revisao(resultado, token, aviso, downloads=None, pasta_saida='', pendencias=None,
                 info=None, edicoes=None, avisos=None, lote_id=None, previa=None,
                 previa_avisos=None):
    # Contagem por SUBCLASSE (igual ao detalhamento abaixo e ao repasse) — antes o
    # topo somava Cirurgias + Procedimentos num número só ("Cirurgias" inflado) e
    # mostrava o total geral como "Procedimentos". (Diretoria 2026-06-24.)
    cont = Counter(medplus.subclasse_preview(p)
                   for b in resultado.blocos for p in b.procedimentos
                   if p.status_calculo != 'componente')
    qtd_cirurgias = cont.get(medplus.SUBCLASSE_CIRURGIA, 0)
    # Todas as subclasses presentes (somam o total) — reconciliam com o detalhamento.
    contagem_subclasses = [(s, cont[s]) for s in medplus.SUBCLASSES if cont.get(s)]
    return {
        'resultado': resultado,
        'token': token,
        'aviso_regras': aviso,
        'pendencias': pendencias if pendencias is not None else _pendencias_revisao(resultado),
        'info': info or [],
        # avisos do chamador (duplicidade) + avisos do PARSER (linhas com dados
        # ignoradas por data ilegível) — nada some em silêncio. (Auditoria 2026-07-02.)
        'avisos': (avisos or []) + list(getattr(resultado, 'avisos', [])),
        'edicoes': edicoes or {},
        'downloads': downloads or [],
        'previa': previa or [],
        'previa_avisos': previa_avisos or [],
        'pasta_saida': pasta_saida,
        'lote_id': lote_id,
        'qtd_cirurgias': qtd_cirurgias,
        'contagem_subclasses': contagem_subclasses,
        'classes': medplus.CLASSES,
        # Médicos novos (sem cadastro) a classificar + as categorias disponíveis.
        'medicos_novos': getattr(resultado, 'medicos_novos', []),
        'categorias_medico': Medico.CATEGORIA_CHOICES,
        # Médicos (menos residentes) p/ o seletor de "memorizar para quais médicos".
        'medicos_memo': list(Medico.objects.exclude(categoria=Medico.CATEGORIA_RESIDENTE)
                             .order_by('nome').values_list('nome', flat=True)),
        'fellows': list(Medico.objects.filter(eh_fellow=True)),
        # Todos os anestesistas (inclui Dra. Regina e Equipe Dra. Regina). A Regina,
        # se escolhida, vira lembrete em dinheiro em vez de repasse.
        'anestesistas': list(Medico.objects.filter(eh_anestesista=True).order_by('nome')),
        'lembretes_regina': getattr(resultado, 'lembretes_regina', []),
        'classe_indefinida': medplus.CLASSE_INDEFINIDA,
        # Preceptores que podem receber a agenda "Equipe Dr. Keiti" + sentinela "Sem preceptor".
        'equipe_destinos': _equipe_destinos(),
        'sem_preceptor': _SEM_PRECEPTOR,
        # Aviso persistente: dias úteis anteriores sem repasse exportado (lacunas).
        'dias_faltantes': _dias_faltantes_fmt(),
    }


def exportar(request):
    """Passo 2: gera os arquivos a partir do relatório já revisado."""
    if request.method != 'POST':
        raise Http404()
    token = request.POST.get('token', '')
    caminho = _caminho_upload(token)
    if caminho is None:
        messages.error(request, 'Arquivo da importação não encontrado — refaça o upload.')
        return redirect('repasses:importar')

    _salvar_rascunho(token, request.POST)               # persiste as edições
    # _preparar_revisao já resolve anestesistas, catarata E preceptoria
    resultado, aviso, dados = _preparar_revisao(token)

    # Campos obrigatórios a verificar (classe/forma de pagamento/fellow/anestesista):
    # bloqueia a exportação e devolve a revisão com os campos destacados p/ piscar.
    pend_obrig = _campos_pendentes(resultado, request.POST)
    if pend_obrig:
        ctx = _ctx_revisao(resultado, token, aviso, edicoes=dados,
                           pendencias=[d for _, d in pend_obrig])
        ctx['campos_pendentes'] = [c for c, _ in pend_obrig]
        ctx['info'] = ['Confirme os campos destacados (em amarelo) antes de exportar — '
                       'inclusive os "sem anestesista"/"sem fellow", para não passar nada batido.']
        return render(request, 'repasses/revisao.html', ctx)

    # Memoriza correções e classes ANTES de qualquer remoção de duplicados: a correção
    # de honorário que a diretoria impôs é uma regra POR procedimento/médico e deve ser
    # gravada mesmo que ESTA linha seja duplicada e removida, ou que o lote inteiro vire
    # "nada novo". memorizar é idempotente (update_or_create). Corrige a regressão em que
    # as correções sumiam ao usar "remover duplicados". (Diretoria 2026-07-01.)
    info = _memorizar_correcoes(resultado, request.POST)
    info += _memorizar_classes(resultado, request.POST)

    # Decisões da revisão são FINALIZADAS aqui, ANTES do dedup: OCI de residente integra no
    # Dr. Alessander (ou é descartado) e a Equipe Dr. Keiti "Sem preceptor" sai. Tem de vir
    # ANTES da remoção de duplicados para o dedup/fingerprint verem o estado final (OCI já
    # no Alessander) e NUNCA removerem um OCI antes de integrá-lo — senão perde pagamento.
    # (Mesma lição do fix das correções.) (Diretoria 2026-07-01.)
    _aplicar_oci_residentes(resultado)
    resultado.blocos = [b for b in resultado.blocos
                        if not getattr(b, 'equipe_sem_preceptor', False)]

    # Lançamentos DUPLICADOS: atendimentos idênticos (mesmo dia/médico/paciente/proc/
    # convênio/VALOR) já exportados em outro lote. Em vez de travar, oferece REMOVER só
    # os duplicados (exporta o que é novo) OU exportar tudo. (Diretoria 2026-06-30.)
    remover_dup = request.POST.get('remover_duplicados') == '1'
    n_removidos = 0
    if remover_dup:
        n_removidos = _remover_duplicados(token, resultado)
        if not (any(repasse.pagaveis(b) for b in resultado.blocos) or resultado.anestesistas):
            # Re-monta o resultado ORIGINAL para o formulário: o atual já foi FINALIZADO
            # e mutado (OCI movido, equipe removida, duplicados fora) — renderizá-lo
            # corromperia os índices posicionais dos campos. (Auditoria 2026-07-02.)
            resultado_form, aviso_form, dados_form = _preparar_revisao(token)
            ctx = _ctx_revisao(resultado_form, token, aviso_form, edicoes=dados_form)
            ctx['info'] = [f'Todos os {n_removidos} lançamento(s) já tinham sido exportados antes '
                           '— nada novo para exportar neste arquivo (as correções marcadas '
                           'foram memorizadas mesmo assim).']
            return render(request, 'repasses/revisao.html', ctx)
    elif request.POST.get('forcar_duplicado') != '1':
        dups = _atendimentos_duplicados(token, resultado)
        if dups:
            # o resultado já foi FINALIZADO (blocos de residente/equipe removidos); para a
            # tela da oferta manter o formulário intacto (mesmos índices dos campos),
            # re-monta o resultado ORIGINAL da revisão só para renderizar.
            resultado_form, aviso_form, dados_form = _preparar_revisao(token)
            ctx = _ctx_revisao(resultado_form, token, aviso_form, edicoes=dados_form)
            ctx['dup_atendimentos'] = [
                f'{b.profissional} — {p.procedimento[:40]} '
                f'({p.data_curta or p.data_texto}; R$ {p.honorario:.2f})' for b, p in dups[:40]]
            ctx['dup_total'] = len(dups)
            ctx['info'] = [f'{len(dups)} lançamento(s) já foram exportados antes (idênticos). '
                           'Remova-os desta exportação ou confirme exportar tudo.']
            return render(request, 'repasses/revisao.html', ctx)

    if n_removidos:
        info.insert(0, f'{n_removidos} lançamento(s) duplicado(s) removido(s) desta exportação '
                       '(já tinham saído em outro lote).')

    # ---- UM LOTE POR DIA (diretoria 2026-07-02) --------------------------------
    # O resultado revisado é dividido pelos dias do atendimento: cada dia vira um
    # lote próprio (token '<upload>~AAAA-MM-DD') com seus arquivos (PDF/Excel por
    # médico + OMIE do dia). Assim um dia pode ser excluído/reaberto sem mexer nos
    # outros; a saída de VÁRIOS dias juntos sai pela seleção de dias no histórico.
    ref = omie._data_referencia(resultado)
    dias_export = sorted({(b.data or ref) for b in resultado.blocos}
                         | {(a.get('data') or ref) for a in resultado.anestesistas})

    # Status de pagamento do formato antigo (um lote por período): preserva ao migrar.
    legado = Lote.objects.filter(token=token).first()
    status_antigos = {}
    if legado:
        status_antigos = {(r.tipo, r.medico, r.data, r.clinica): r.status
                          for r in legado.repasses.exclude(status=Repasse.STATUS_GERADO)}

    por_dia, pend = [], []
    for dia in dias_export:
        res_dia = _resultado_do_dia(resultado, dia, ref)
        if not (any(repasse.pagaveis(b) for b in res_dia.blocos) or res_dia.anestesistas):
            continue                     # dia sem nada pagável não vira lote
        arqs_dia, pend_dia = _gerar_arquivos_por_dia(res_dia)
        pend += pend_dia
        por_dia.append((dia, res_dia, arqs_dia))

    todos_arquivos = [a for _, _, arqs in por_dia for a in arqs]
    pasta_saida, downloads = _salvar_saidas(todos_arquivos)

    # Persistência ATÔMICA: lotes-dia + arquivos + repasses + migração/limpeza do legado
    # entram juntos ou nada — sem lote meio-gravado se algo falhar no meio.
    # (Auditoria 2026-07-02.)
    from django.db import transaction
    lotes_dia, fps_todos = [], []
    with transaction.atomic():
        for dia, res_dia, arqs in por_dia:
            dls = [{'grupo': g, 'arquivo': n} for (g, n, _c) in arqs]
            lote_dia, fps = _registrar_lote(request, token, f'{token}~{dia.isoformat()}',
                                            res_dia, dados, pasta_saida, dls)
            _guardar_arquivos_no_banco(lote_dia, arqs)   # re-download não depende da pasta saídas/
            _guardar_upload_no_banco(lote_dia)           # re-export sobrevive à limpeza de uploads/
            fps_todos += fps
            lotes_dia.append(lote_dia)

        # Migra os status do lote legado para os lotes-dia e limpa o que sobrou: o legado
        # e lotes-dia de dias que não existem mais nesta re-exportação.
        if status_antigos:
            for l in lotes_dia:
                for r in l.repasses.filter(status=Repasse.STATUS_GERADO):
                    st = status_antigos.get((r.tipo, r.medico, r.data, r.clinica))
                    if st:
                        r.status = st
                        r.save(update_fields=['status'])
        (Lote.objects.filter(token__startswith=token)
         .exclude(token__in={l.token for l in lotes_dia}).delete())

    avisos_dup = _avisos_duplicidade(token, fps_todos)
    # Formulário pós-exportação: re-monta o resultado ORIGINAL da revisão — o atual foi
    # mutado (OCI movido, equipe/dups removidos) e os índices dos campos deixariam de
    # casar se a pessoa editasse e exportasse de novo. Antes de apagar o rascunho.
    # (Auditoria 2026-07-02.)
    resultado_form, aviso_form, dados_form = _preparar_revisao(token)
    # Exportado -> larga o cache de edições (vivem no lote agora; reabrir pelo histórico).
    # O "Continuar edição" deixa de oferecer este import. (Diretoria 2026-06-24.)
    RepasseRascunho.objects.filter(token=token).delete()
    pendencias = _resumir_pendencias(pend)
    if len(lotes_dia) > 1:
        rot = ', '.join(l.periodo_inicio.strftime('%d/%m') for l in lotes_dia if l.periodo_inicio)
        info.append(f'Exportação registrada em {len(lotes_dia)} lotes — um por dia ({rot}). '
                    'Cada dia pode ser excluído ou reaberto separadamente no histórico.')
    ctx = _ctx_revisao(resultado_form, token, aviso_form, downloads, pasta_saida, pendencias,
                       info=info, edicoes=dados_form, avisos=avisos_dup,
                       lote_id=lotes_dia[0].id if lotes_dia else None)
    return render(request, 'repasses/revisao.html', ctx)


def _resultado_do_dia(resultado, dia, ref):
    """Cópia rasa do resultado só com os blocos/anestesistas DO DIA — para gerar os
    arquivos e registrar o lote daquele dia. Blocos sem data entram no dia de
    referência (mesma regra de omie.linhas_relatorio_pagar)."""
    import copy
    res = copy.copy(resultado)
    res.blocos = [b for b in resultado.blocos if (b.data or ref) == dia]
    res.anestesistas = [a for a in resultado.anestesistas if (a.get('data') or ref) == dia]
    return res


def _memorizar_correcoes(resultado, post):
    """Salva como correção memorizada cada linha marcada 'memorizar' na revisão.

    A correção é POR MÉDICO: vale só para o médico do bloco e para os que o usuário
    marcou no seletor (modal). Assim mudar o valor de um procedimento do Dr. Rodolpho
    NÃO altera o mesmo procedimento da Dra. Tharcila — a menos que ela seja marcada.
    Devolve mensagens informativas."""
    salvos = 0
    for bloco in resultado.blocos:
        # Nome CANÔNICO do médico do bloco (do cadastro, resolvido em processar) —
        # mesma chave que o aplicar() usa, para a correção casar nos próximos meses
        # mesmo com o nome do MedPlus diferente (sufixo/abreviação). Ver correcoes.medico_norm.
        dono = getattr(bloco, 'medico_cadastro', '') or bloco.profissional
        for p in bloco.procedimentos:
            if post.get(f'memorizar_{p.idx}') != '1':
                continue
            if p.honorario is None or p.honorario <= 0:
                continue
            # médicos que recebem a correção: o do bloco + os marcados no modal (nomes do cadastro).
            destinos = [d.strip() for d in post.getlist(f'memo_medicos_{p.idx}') if d.strip()]
            if dono not in destinos:
                destinos.insert(0, dono)
            for medico in dict.fromkeys(destinos):   # dedup, preserva ordem
                correcoes.memorizar(
                    p.procedimento, p.convenio, round(float(p.honorario), 2),
                    medico=medico, classe=p.classe,
                    origem=(resultado.unidade or '')[:180],
                    observacao=f'{bloco.profissional} {p.data_texto}'.strip())
            salvos += 1
    if not salvos:
        return []
    plural = 'correção memorizada' if salvos == 1 else 'correções memorizadas'
    return [f'✓ {salvos} {plural} (por médico) — serão reaplicadas automaticamente nos '
            'próximos meses. Veja em "Correções memorizadas".']


def _memorizar_classes(resultado, post):
    """Memoriza a CLASSE de um procedimento quando a diretoria a define/muda na
    revisão. A classe é intrínseca ao procedimento, então vale para TODOS os médicos
    e é reaplicada em todos os lançamentos futuros desse procedimento. (Diretoria
    2026-06-26.) Devolve mensagens informativas."""
    salvos = set()
    for bloco in resultado.blocos:
        for p in bloco.procedimentos:
            # Pula sintéticas, status com fluxo próprio (catarata/componente) e taxas
            # de sala (classe vem do convênio) — alinhado com correcoes.aplicar_classes.
            if (getattr(p, 'sintetica', False)
                    or getattr(p, 'status_calculo', None) in correcoes._NAO_SOBRESCREVER
                    or 'taxa' in regras.normalizar(p.convenio)):
                continue
            sug = getattr(p, 'classe_sugerida', None)
            classe = (p.classe or '').strip()
            # Só memoriza quando o usuário MUDA a classe (difere da sugerida).
            if not (sug is not None and classe and classe != medplus.CLASSE_INDEFINIDA
                    and classe != sug):
                continue
            chave = regras.normalizar(p.procedimento)
            if not chave or chave in salvos:
                continue
            # NÃO re-memoriza nem REATIVA uma classe já guardada com o MESMO valor:
            # reabrir/re-exportar um lote antigo não deve mexer na memória nem
            # ressuscitar uma classe que a diretoria desligou. (Diretoria 2026-06-26.)
            ja = ClasseMemorizada.objects.filter(proc_norm=chave).first()
            if ja is not None and ja.classe == classe:
                continue
            correcoes.memorizar_classe(p.procedimento, classe,
                                       origem=(resultado.unidade or '')[:180])
            salvos.add(chave)
    if not salvos:
        return []
    n = len(salvos)
    plural = 'classe memorizada' if n == 1 else 'classes memorizadas'
    return [f'✓ {n} {plural} por procedimento — a mesma classe será aplicada '
            'automaticamente nos próximos lançamentos desse procedimento.']


GRUPO_OMIE = 'Importação OMIE'   # subpasta/rotulagem dos arquivos OMIE (a pagar / a receber)
_MESES_PT = ('janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho',
             'agosto', 'setembro', 'outubro', 'novembro', 'dezembro')


def _periodo_resultado(resultado):
    """(menor, maior) data dos atendimentos do resultado; (None, None) se não houver."""
    datas = [p.data for b in resultado.blocos for p in b.procedimentos if getattr(p, 'data', None)]
    return (min(datas), max(datas)) if datas else (None, None)


def _rotulo_periodo_dias(ini, fim):
    """Faixa de dias p/ nome de zip: '27-29' (mesmo mês), '27' (1 dia),
    '27-06_a_02-07' (meses diferentes). Vazio quando não há datas."""
    if not ini or not fim:
        return ''
    if ini == fim:
        return f'{ini.day:02d}'
    if (ini.year, ini.month) == (fim.year, fim.month):
        return f'{ini.day:02d}-{fim.day:02d}'
    return f'{ini.day:02d}-{ini.month:02d}_a_{fim.day:02d}-{fim.month:02d}'


def _rotulo_periodo_extenso(ini, fim):
    """Faixa p/ nome de arquivo OMIE: '27-29 de Junho', '27 de Junho',
    '27 de Junho a 02 de Julho'. Vazio quando não há datas."""
    if not ini or not fim:
        return ''
    mes_ini = _MESES_PT[ini.month - 1].capitalize()
    if ini == fim:
        return f'{ini.day:02d} de {mes_ini}'
    if (ini.year, ini.month) == (fim.year, fim.month):
        return f'{ini.day:02d}-{fim.day:02d} de {mes_ini}'
    mes_fim = _MESES_PT[fim.month - 1].capitalize()
    return f'{ini.day:02d} de {mes_ini} a {fim.day:02d} de {mes_fim}'


def _nome_omie(base, ini, fim):
    """'OMIE_Contas_a_Pagar_27-29 de Junho.xlsx' (sem sufixo se não houver datas)."""
    suf = _rotulo_periodo_extenso(ini, fim)
    return f'{base}_{suf}.xlsx' if suf else f'{base}.xlsx'


def _previa_arquivos(resultado):
    """(grupo, arquivo) que a exportação vai gerar — só os NOMES (sem gerar bytes),
    p/ a pessoa conferir após Salvar quais médicos têm repasse. Mesma nomenclatura
    de _gerar_arquivos_por_dia — manter os dois em sincronia."""
    ini, fim = _periodo_resultado(resultado)
    out = [{'grupo': GRUPO_OMIE, 'arquivo': _nome_omie('OMIE_Contas_a_Pagar', ini, fim)},
           {'grupo': GRUPO_OMIE, 'arquivo': _nome_omie('OMIE_Contas_a_Receber', ini, fim)}]
    for bloco in resultado.blocos:
        if not repasse.pagaveis(bloco):
            continue
        base = repasse.nome_base(bloco)
        grupo = f'Repasse — {bloco.profissional}'
        out.append({'grupo': grupo, 'arquivo': f'{base}.xlsx'})
        out.append({'grupo': grupo, 'arquivo': f'{base}.pdf'})
    for a in resultado.anestesistas:
        base = repasse.nome_base_anestesista(a)
        grupo = f'Anestesista — {a["anestesista"]}'
        out.append({'grupo': grupo, 'arquivo': f'{base}.xlsx'})
        out.append({'grupo': grupo, 'arquivo': f'{base}.pdf'})
    return out


def _gerar_arquivos_por_dia(resultado):
    """Gera os arquivos de saída.

    OMIE: UM contas a pagar + UM contas a receber para todo o período — cada
    linha já carrega sua própria data (registro/vencimento) e a observação com a
    data abreviada, então não há mais necessidade de separar os arquivos por dia.
    O a pagar sai com uma linha por (médico × dia × clínica × classe) e o a
    receber com uma linha por (dia × clínica).

    Repasses (Excel + PDF): um por bloco, ou seja, por médico/dia/clínica.
    """
    arquivos, pend = [], []

    ini, fim = _periodo_resultado(resultado)
    pagar = omie.gerar_contas_pagar(resultado, settings.OMIE_PAGAR_TEMPLATE)
    receber = omie.gerar_contas_receber(resultado, settings.OMIE_RECEBER_TEMPLATE)
    arquivos.append((GRUPO_OMIE, _nome_omie('OMIE_Contas_a_Pagar', ini, fim), pagar.conteudo))
    arquivos.append((GRUPO_OMIE, _nome_omie('OMIE_Contas_a_Receber', ini, fim), receber.conteudo))
    pend += pagar.pendencias + receber.pendencias

    for bloco in resultado.blocos:
        if not repasse.pagaveis(bloco):
            continue
        base = repasse.nome_base(bloco)
        grupo = f'Repasse — {bloco.profissional}'
        arquivos.append((grupo, f'{base}.xlsx', repasse.gerar_excel(bloco, resultado.unidade)))
        arquivos.append((grupo, f'{base}.pdf', repasse.gerar_pdf(bloco, resultado.unidade)))
    for a in resultado.anestesistas:
        base = repasse.nome_base_anestesista(a)
        grupo = f'Anestesista — {a["anestesista"]}'
        arquivos.append((grupo, f'{base}.xlsx', repasse.gerar_excel_anestesista(a, resultado.unidade)))
        arquivos.append((grupo, f'{base}.pdf', repasse.gerar_pdf_anestesista(a, resultado.unidade)))
    return arquivos, pend


def baixar_saida(request, pasta, arquivo):
    """Serve um arquivo gerado da pasta de saídas (com proteção contra path traversal)."""
    base = Path(settings.SAIDAS_DIR).resolve()
    caminho = (base / pasta / arquivo).resolve()
    if base not in caminho.parents or not caminho.is_file():
        raise Http404('Arquivo não encontrado.')
    return FileResponse(open(caminho, 'rb'), as_attachment=True, filename=arquivo)


def _guardar_arquivos_no_banco(lote, arquivos):
    """Guarda os bytes dos arquivos gerados no banco, ligados ao lote(-dia), para
    re-download mesmo se a pasta saídas/ for limpa. Re-export substitui."""
    if not lote:
        return
    lote.arquivos.all().delete()
    ArquivoSaida.objects.bulk_create([
        ArquivoSaida(lote=lote, grupo=g, nome=n, conteudo=conteudo)
        for (g, n, conteudo) in arquivos
    ])


def _guardar_upload_no_banco(lote):
    """Guarda o .xls de origem no lote(-dia) (uma vez), para re-exportar/reabrir mesmo
    que a pasta uploads/ seja limpa — cada lote-dia carrega o arquivo inteiro, então
    reabrir funciona mesmo que os outros dias tenham sido excluídos."""
    if not lote or lote.upload_conteudo:
        return
    caminho = _caminho_upload(lote.token)   # resolve o token-base internamente
    if caminho and caminho.is_file():
        lote.upload_conteudo = caminho.read_bytes()
        lote.save(update_fields=['upload_conteudo'])


def baixar_lote(request, pk, nome):
    """Re-download de um arquivo do histórico — servido do BANCO (independe do disco)."""
    arq = ArquivoSaida.objects.filter(lote_id=pk, nome=nome).first()
    if arq is None:
        raise Http404('Arquivo não encontrado.')
    return FileResponse(io.BytesIO(bytes(arq.conteudo)), as_attachment=True, filename=nome)


def baixar_lote_zip(request, pk):
    """Baixa TODOS os arquivos do lote de uma vez, num único .zip — para validar sem
    baixar arquivo por arquivo. Servido do BANCO (ArquivoSaida)."""
    import zipfile
    lote = get_object_or_404(Lote, pk=pk)
    arquivos = list(lote.arquivos.all())
    if not arquivos:
        raise Http404('Este lote não tem arquivos guardados para baixar.')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        usados = {}
        for a in arquivos:
            # Agrupa por TIPO de arquivo (não por médico): OMIE / PDF / XLSX —
            # a diretoria separa os PDFs, as planilhas e os arquivos OMIE. (2026-06-30.)
            ext = a.nome.rpartition('.')[2].lower()
            if a.nome.startswith('OMIE') or (a.grupo or '') == GRUPO_OMIE:
                pasta = 'OMIE'
            elif ext == 'pdf':
                pasta = 'PDF'
            elif ext == 'xlsx':
                pasta = 'XLSX'
            else:
                pasta = 'Outros'
            nome = f'{pasta}/{a.nome}'
            usados[nome] = usados.get(nome, 0) + 1
            if usados[nome] > 1:
                base, _, ext2 = a.nome.rpartition('.')
                nome = f'{pasta}/{base} ({usados[nome]}).{ext2}' if ext2 else f'{pasta}/{a.nome} ({usados[nome]})'
            zf.writestr(nome, bytes(a.conteudo))
    buf.seek(0)
    # Nome do zip = "Repasses_Médicos_<faixa> de <Mês>" (com o mês, igual aos arquivos
    # OMIE), ex.: "Repasses_Médicos_27-29 de Junho". (Diretoria 2026-06-30.)
    faixa = _rotulo_periodo_extenso(lote.periodo_inicio, lote.periodo_fim)
    nome_zip = f'Repasses_Médicos_{faixa}.zip' if faixa else 'Repasses_Médicos.zip'
    return FileResponse(buf, as_attachment=True, filename=nome_zip)


# --- Histórico de lotes + auditoria -------------------------------------------

def _arquivo_nome(token):
    r = RepasseRascunho.objects.filter(token=token).first()
    return (r.arquivo_nome if r else '') or ''


def _fingerprint(p, profissional):
    """Impressão digital de um atendimento, para detectar duplicidade entre lotes.

    Inclui o PACIENTE (discrimina atendimentos iguais de pacientes diferentes no
    mesmo dia) e NÃO inclui a clínica (um export pode trazer a coluna Clínica e
    outro não — o mesmo atendimento tem que casar mesmo assim)."""
    return '|'.join([
        p.data.isoformat() if p.data else '',
        regras.normalizar(profissional),
        regras.normalizar(getattr(p, 'paciente', '') or ''),
        regras.normalizar(p.procedimento),
        regras.normalizar(p.convenio or ''),
        ('%.2f' % float(p.valor)) if p.valor is not None else '',
    ])


def _fingerprints_pagaveis(resultado):
    """Impressões digitais dos atendimentos PAGÁVEIS deste resultado (p/ duplicidade)."""
    fps = []
    for b in resultado.blocos:
        for p in b.procedimentos:
            if p.status_calculo == 'calculado' and (p.honorario or 0) > 0:
                fps.append(_fingerprint(p, b.profissional))
    return fps


def _fingerprints_outros_lotes(token):
    """Multiconjunto (Counter) das impressões digitais já exportadas em OUTROS lotes
    (qualquer upload diferente deste token — a família inteira de lotes-dia deste
    upload é excluída). O fingerprint inclui a data com ANO e o VALOR, então só casa
    atendimento idêntico, do mesmo ano e mesmo valor."""
    prior = Counter()
    for outro in (Lote.objects.exclude(token__startswith=_token_base(token))
                  .only('fingerprints')):
        prior.update(outro.fingerprints or [])
    return prior


def _atendimentos_duplicados(token, resultado):
    """[(bloco, procedimento)] dos atendimentos PAGÁVEIS deste resultado que já saíram
    IDÊNTICOS (mesmo fingerprint = mesmo dia/médico/paciente/procedimento/convênio/valor)
    em outro lote. Multiconjunto: se o arquivo traz 2 iguais e só 1 já saiu, marca 1."""
    prior = _fingerprints_outros_lotes(token)
    if not prior:
        return []
    usados, dups = Counter(), []
    for b in resultado.blocos:
        for p in b.procedimentos:
            if not (p.status_calculo == 'calculado' and (p.honorario or 0) > 0):
                continue
            fp = _fingerprint(p, b.profissional)
            if usados[fp] < prior.get(fp, 0):
                usados[fp] += 1
                dups.append((b, p))
    return dups


def _remover_duplicados(token, resultado):
    """Remove do resultado os atendimentos idênticos já exportados antes (e os
    anestesistas cujas cirurgias saíram todas). Devolve quantos atendimentos removeu."""
    dups = _atendimentos_duplicados(token, resultado)
    if not dups:
        return 0
    remover = {id(p) for _, p in dups}
    for b in resultado.blocos:
        b.procedimentos = [p for p in b.procedimentos if id(p) not in remover]
    resultado.blocos = [b for b in resultado.blocos if b.procedimentos]
    vivos = {id(p) for b in resultado.blocos for p in b.procedimentos}
    resultado.anestesistas = [a for a in getattr(resultado, 'anestesistas', [])
                              if not a.get('cirurgias')
                              or any(id(c) in vivos for c in a['cirurgias'])]
    return len(dups)


def _auditoria_lote(resultado, dados):
    """Lista legível dos ajustes manuais — a trilha de auditoria do lote."""
    itens = list(getattr(resultado, 'log_edicoes', []))
    for a in resultado.anestesistas:
        extra = f' (+{a["horas"]}h extra)' if a.get('horas') else ''
        itens.append(f'Anestesista: {a["anestesista"]} em {a["cirurgiao"]}{extra} — '
                     + ('R$ %.2f' % float(a.get('valor') or 0)))
    for chave, val in dados.items():
        if chave.startswith('cat_modo_') and val:
            itens.append(f'Catarata definida: {val}.')
        elif chave.startswith('preceptoria_') and val:
            itens.append(f'Preceptoria informada: R$ {val}.')
    return itens


def _avisos_duplicidade(token_base, fps):
    """Avisos de duplicidade: atendimentos que já apareceram em lote de OUTRO upload.
    A família inteira de lotes-dia deste upload é excluída (token__startswith); NÃO
    pulamos por nome de arquivo (dois uploads podem ter o mesmo nome). Conta por
    multiconjunto para não subnotificar atendimentos repetidos."""
    avisos, novos = [], Counter(fps)
    # só os campos lidos — NÃO arrasta o upload_conteudo (BinaryField pesado).
    outros = (Lote.objects.exclude(token__startswith=token_base)
              .only('id', 'arquivo_nome', 'criado_em', 'fingerprints'))
    for outro in outros:
        inter = novos & Counter(outro.fingerprints or [])
        n = sum(inter.values())
        if n:
            avisos.append(f'{n} atendimento(s) já saíram no lote #{outro.id} '
                          f'({outro.arquivo_nome or "?"}, {outro.criado_em:%d/%m/%Y}) — '
                          'confira para não pagar 2×.')
    return avisos


def _registrar_lote(request, token_base, token_lote, resultado, dados, pasta_saida, downloads):
    """Cria/atualiza UM lote (de um dia) da exportação e devolve suas fingerprints.
    `resultado` aqui já vem filtrado para o dia do lote."""
    fps, medicos, n_repasses = [], set(), 0
    total_pagar = 0.0
    for b in resultado.blocos:
        pagaveis = [p for p in b.procedimentos
                    if p.status_calculo == 'calculado' and (p.honorario or 0) > 0]
        if pagaveis:
            medicos.add(regras.normalizar(b.profissional))
            n_repasses += 1
        for p in pagaveis:
            total_pagar += p.honorario
            fps.append(_fingerprint(p, b.profissional))
    for a in resultado.anestesistas:
        total_pagar += a.get('valor') or 0
        n_repasses += 1
    total_receber = sum(float(p.valor) for b in resultado.blocos for p in b.procedimentos
                        if p.valor is not None)
    datas = [p.data for b in resultado.blocos for p in b.procedimentos if p.data]
    arq_nome = _arquivo_nome(token_base)
    quem = request.user.get_username() if request.user.is_authenticated else 'diretoria'

    defaults = {
        'criado_por': quem, 'unidade': resultado.unidade or '',
        'periodo_inicio': min(datas) if datas else None,
        'periodo_fim': max(datas) if datas else None,
        'n_medicos': len(medicos), 'n_repasses': n_repasses,
        'total_pagar': round(total_pagar, 2), 'total_receber': round(total_receber, 2),
        'pasta_saida': pasta_saida, 'downloads': downloads,
        'auditoria': _auditoria_lote(resultado, dados), 'fingerprints': fps,
        'edicoes': dados or {},   # rascunho da revisão — permite reabrir e editar o lote
        'linhas_pagar': omie.linhas_relatorio_pagar(resultado),   # p/ o relatório mensal
    }
    if arq_nome:
        defaults['arquivo_nome'] = arq_nome  # não sobrescreve o nome bom com vazio
    lote, _ = Lote.objects.update_or_create(token=token_lote, defaults=defaults)
    _sync_repasses(lote, resultado)
    return lote, fps


def _sync_repasses(lote, resultado):
    """Cria/atualiza os repasses individuais do lote (p/ acompanhar o pagamento).
    Preserva o status de quem já existia (re-exportar não volta tudo p/ 'gerado')."""
    atuais = set()
    for b in resultado.blocos:
        pag = [p for p in b.procedimentos
               if p.status_calculo == 'calculado' and (p.honorario or 0) > 0]
        if not pag:
            continue
        valor = round(sum(p.honorario for p in pag), 2)
        Repasse.objects.update_or_create(
            lote=lote, tipo='medico', medico=b.profissional, data=b.data,
            clinica=b.clinica or '',
            defaults={'valor': valor, 'razao_social': b.razao_social or ''})
        atuais.add(('medico', b.profissional, b.data, b.clinica or ''))
    # Agrega por (anestesista, dia, clínica) ANTES de gravar: a anestesista que atende
    # DOIS cirurgiões no mesmo dia/clínica tem os valores SOMADOS — o update_or_create
    # por chave sobrescrevia o primeiro. (Auditoria 2026-07-02.)
    anest_agg = {}
    for a in resultado.anestesistas:
        chave = (a['anestesista'], a.get('data'), a.get('clinica', '') or '')
        agg = anest_agg.setdefault(chave, {'valor': 0.0, 'razao_social': ''})
        agg['valor'] += float(a.get('valor') or 0)
        agg['razao_social'] = agg['razao_social'] or (a.get('razao_social', '') or '')
    for (nome, data, clinica), agg in anest_agg.items():
        Repasse.objects.update_or_create(
            lote=lote, tipo='anestesista', medico=nome, data=data, clinica=clinica,
            defaults={'valor': round(agg['valor'], 2), 'razao_social': agg['razao_social']})
        atuais.add(('anestesista', nome, data, clinica))
    for r in lote.repasses.all():
        if (r.tipo, r.medico, r.data, r.clinica) not in atuais:
            r.delete()


def _dias_sem_repasse():
    """Dias ÚTEIS (não-domingo) anteriores a HOJE, dentro do período coberto (últimos
    ~60 dias), que ainda NÃO tiveram repasse exportado. Ex.: se hoje é 01/07 e foram
    lançados 29 e 30/06 mas o último antes disso foi 25/06, avisa 26 e 27/06 (o 28/06 é
    domingo, não conta). Serve de aviso persistente no histórico e na revisão.
    (Diretoria 2026-07-01.)"""
    from datetime import date, timedelta
    dias = set(Repasse.objects.exclude(data__isnull=True)
               .values_list('data', flat=True).distinct())
    # dias que a diretoria CONFIRMOU não terem repasse contam como "cobertos".
    dias |= set(DiaSemRepasse.objects.values_list('data', flat=True))
    if not dias:
        return []
    hoje = date.today()
    inicio = max(min(dias), hoje - timedelta(days=60))   # limita a lacunas recentes
    faltantes, d = [], inicio
    while d < hoje:
        if d.weekday() != 6 and d not in dias:           # 6 = domingo (sem atendimento)
            faltantes.append(d)
        d += timedelta(days=1)
    return faltantes


def _dias_faltantes_fmt():
    """Dias sem repasse exportado: [{iso, label}] — para o aviso persistente e o botão
    de confirmar 'não houve repasse'."""
    return [{'iso': d.isoformat(), 'label': d.strftime('%d/%m/%Y')} for d in _dias_sem_repasse()]


def confirmar_sem_repasse(request):
    """A diretoria confirma que um dia NÃO teve repasse (não havia o que exportar). Cria o
    marcador — o dia some do aviso e aparece no histórico com valor nulo. (2026-07-02.)"""
    if request.method != 'POST':
        raise Http404()
    from datetime import date
    try:
        dia = date.fromisoformat((request.POST.get('data') or '').strip())
    except ValueError:
        raise Http404('Data inválida.')
    quem = request.user.get_username() if request.user.is_authenticated else 'diretoria'
    DiaSemRepasse.objects.get_or_create(data=dia, defaults={'criado_por': quem})
    messages.success(request, f'Dia {dia:%d/%m/%Y} confirmado sem repasse.')
    return redirect('repasses:lotes')


def remover_sem_repasse(request, pk):
    """Remove a confirmação de 'dia sem repasse' — o dia volta a ser solicitado no aviso."""
    if request.method != 'POST':
        raise Http404()
    obj = DiaSemRepasse.objects.filter(pk=pk).first()
    if obj:
        dia = obj.data
        obj.delete()
        messages.info(request, f'Dia {dia:%d/%m/%Y} voltou a ser solicitado (confirmação removida).')
    return redirect('repasses:lotes')


_MESES_FILTRO = [(1, 'Janeiro'), (2, 'Fevereiro'), (3, 'Março'), (4, 'Abril'), (5, 'Maio'),
                 (6, 'Junho'), (7, 'Julho'), (8, 'Agosto'), (9, 'Setembro'), (10, 'Outubro'),
                 (11, 'Novembro'), (12, 'Dezembro')]


def _dias_do_historico():
    """Um LANÇAMENTO POR DIA (diretoria 2026-07-02): agrupa os repasses de todos os
    lotes pelo dia do atendimento — lançar 12–20/06 vira uma linha para cada dia com
    movimento (12, 13, 15, ... — domingo sem atendimento não aparece). Os dias
    confirmados como "sem repasse" entram com valor nulo (R$ 0,00). Uma query só."""
    dias = {}
    for data, lote_id, medico, valor, status in (
            Repasse.objects.exclude(data__isnull=True)
            .values_list('data', 'lote_id', 'medico', 'valor', 'status')):
        d = dias.setdefault(data, {'data': data, 'iso': data.isoformat(), 'total': 0,
                                   'medicos': set(), 'lotes': set(), 'n': 0, 'pagos': 0,
                                   'sem_repasse': None})
        d['total'] += valor
        d['medicos'].add(medico)
        d['lotes'].add(lote_id)
        d['n'] += 1
        if status == Repasse.STATUS_PAGO:
            d['pagos'] += 1
    # dias confirmados "sem repasse" = lançamento nulo no histórico (some se deletado)
    for c in DiaSemRepasse.objects.all():
        if c.data not in dias:
            dias[c.data] = {'data': c.data, 'iso': c.data.isoformat(), 'total': 0,
                            'medicos': set(), 'lotes': set(), 'n': 0, 'pagos': 0,
                            'sem_repasse': c}
    out = []
    for d in sorted(dias.values(), key=lambda x: x['data'], reverse=True):
        d['n_medicos'] = len(d['medicos'])
        d['lotes'] = sorted(d['lotes'])
        d['pendentes'] = d['n'] - d['pagos']
        d['excluir_lote'] = None
        out.append(d)
    # Dias cobertos por UM lote-dia (período = o próprio dia) e sem repasse pago podem
    # ser excluídos direto do histórico — o "excluir só um dia" da diretoria (2026-07-02).
    unicos = {d['lotes'][0] for d in out if len(d['lotes']) == 1}
    if unicos:
        info = {l['id']: l for l in Lote.objects.filter(id__in=unicos)
                .annotate(pagos_n=Count('repasses', filter=Q(repasses__status=Repasse.STATUS_PAGO)))
                .values('id', 'periodo_inicio', 'periodo_fim', 'pagos_n')}
        for d in out:
            if len(d['lotes']) != 1:
                continue
            li = info.get(d['lotes'][0])
            if li and li['periodo_inicio'] == li['periodo_fim'] == d['data'] and not li['pagos_n']:
                d['excluir_lote'] = li['id']
    return out


def lotes_lista(request):
    """Histórico com UM LANÇAMENTO POR DIA (+ a lista dos arquivos exportados/lotes).
    Filtros de mês/ano, confirmação de dias sem repasse e seleção de dias para emitir
    relatório específico (conferência do contas a pagar na OMIE)."""
    ano = request.GET.get('ano')
    mes = request.GET.get('mes')
    if ano is None and mes is None:
        # Primeira visita (sem filtro na URL): pré-seleciona o mês/ano ATUAL
        # (diretoria 2026-07-02). Escolher "Todos" envia ano=/mes= vazios e passa.
        from datetime import date as _hoje_cls
        hoje = _hoje_cls.today()
        ano, mes = str(hoje.year), str(hoje.month)
    ano = ano or ''
    mes = mes or ''

    dias_todos = _dias_do_historico()
    dias = [d for d in dias_todos
            if (not ano.isdigit() or d['data'].year == int(ano))
            and (not mes.isdigit() or d['data'].month == int(mes))]

    # Anos p/ o filtro: dos dias lançados E dos períodos dos lotes (cobre as 2 tabelas).
    anos = sorted({d['data'].year for d in dias_todos}
                  | {p.year for p in Lote.objects.exclude(periodo_inicio__isnull=True)
                     .values_list('periodo_inicio', flat=True)}, reverse=True)

    # defer: NÃO arrasta o .xls de origem (upload_conteudo) nem os JSONs pesados de
    # cada lote só para listar — o render usa apenas as colunas exibidas + linhas_pagar
    # (filiais_resumo). (Auditoria 2026-07-02.)
    qs = (Lote.objects
          .defer('upload_conteudo', 'fingerprints', 'edicoes', 'auditoria')
          .annotate(n_total=Count('repasses'),
                    n_pagos=Count('repasses', filter=Q(repasses__status=Repasse.STATUS_PAGO))))
    if ano.isdigit():
        qs = qs.filter(periodo_inicio__year=int(ano))
    if mes.isdigit():
        qs = qs.filter(periodo_inicio__month=int(mes))
    lotes = list(qs)
    for l in lotes:
        l.n_pendentes = l.n_total - l.n_pagos

    return render(request, 'repasses/lotes.html', {
        'dias': dias,
        'lotes': lotes,
        'total_dias': len(dias),
        'total_pagar': sum((d['total'] for d in dias), 0),
        'pendentes_total': sum(d['pendentes'] for d in dias),
        'dias_faltantes': _dias_faltantes_fmt(),
        'dias_confirmados': list(DiaSemRepasse.objects.all()),
        'anos': anos,
        'meses': _MESES_FILTRO,
        'ano_sel': ano,
        'mes_sel': mes,
    })


def relatorio_dias(request):
    """Emite o relatório "Repasses em Aberto" SÓ dos dias selecionados no histórico —
    para conferir/reajustar o contas a pagar na OMIE quando um dia muda.
    (Diretoria 2026-07-02.)"""
    from datetime import date as _date
    from .services import relatorio
    if request.method != 'POST':
        raise Http404()
    dias = []
    for s in request.POST.getlist('dias'):
        try:
            _date.fromisoformat(s)
            dias.append(s)
        except (TypeError, ValueError):
            continue
    dias = sorted(set(dias))
    if not dias:
        messages.error(request, 'Selecione ao menos um dia para emitir o relatório.')
        return redirect('repasses:lotes')
    alvo = set(dias)
    linhas = []
    for l in Lote.objects.only('linhas_pagar'):
        linhas.extend(ln for ln in (l.linhas_pagar or []) if (ln.get('data') or '') in alvo)
    if not linhas:
        messages.error(request, 'Os dias selecionados não têm lançamentos no a pagar.')
        return redirect('repasses:lotes')

    def _fmt(iso):
        return f'{iso[8:10]}/{iso[5:7]}'
    if len(dias) == 1:
        rotulo = f'{_fmt(dias[0])}/{dias[0][:4]}'
    elif len(dias) <= 4:
        rotulo = ', '.join(_fmt(d) for d in dias)
    else:
        rotulo = f'{_fmt(dias[0])} a {_fmt(dias[-1])} ({len(dias)} dias)'
    nome = 'Dias_' + '_'.join(f'{d[8:10]}-{d[5:7]}' for d in dias[:6])
    if len(dias) > 6:
        nome += f'_e_mais_{len(dias) - 6}'

    # formato=omie -> contas a pagar OMIE só dos dias marcados (Fornecedor = CNPJ),
    # p/ reimportar na OMIE quando um dia muda. Padrão: relatório de conferência.
    if request.POST.get('formato') == 'omie':
        livro = regras.carregar_livro_padrao()

        def _fornecedor(nome_med):
            m = livro.medico_por_nome(nome_med) if livro else None
            if m and (m.cnpj or m.razao_social):
                return m.cnpj or m.razao_social
            return nome_med
        linhas_omie = [{'nome': _fornecedor(ln.get('medico') or ''),
                        'categoria': ln.get('categoria') or '',
                        'valor': float(ln.get('valor') or 0),
                        'registro': _date.fromisoformat(ln['data']),
                        'vencimento': (_date.fromisoformat(ln['vencimento'])
                                       if ln.get('vencimento') else _date.fromisoformat(ln['data'])),
                        'departamento': ln.get('departamento') or '',
                        'observacao': f"Repasse {ln.get('medico') or ''} {_fmt(ln['data'])}".strip()}
                       for ln in linhas]
        res = omie.gerar_contas_pagar_de_linhas(linhas_omie, settings.OMIE_PAGAR_TEMPLATE)
        return FileResponse(io.BytesIO(res.conteudo), as_attachment=True,
                            filename=f'OMIE_Contas_a_Pagar_{nome}.xlsx')

    conteudo = relatorio.gerar_relatorio_mensal(linhas, f'Repasses dos dias {rotulo}')
    return FileResponse(io.BytesIO(conteudo), as_attachment=True,
                        filename=f'Repasses_{nome}.xlsx')


def _ultimo_dia_mes(mes):
    """date do último dia do mês 'YYYY-MM' (None se inválido)."""
    import calendar
    from datetime import date
    try:
        ano, m = int(mes[:4]), int(mes[5:7])
        return date(ano, m, calendar.monthrange(ano, m)[1])
    except (ValueError, IndexError):
        return None


def _preceptores_mensais():
    """[(Medico, valor)] dos preceptores com valor MENSAL — cadastro com obs '/mês'
    (não '/semana'). A preceptoria SEMANAL continua sendo lançada na exportação diária;
    a MENSAL é um valor fixo do cadastro, cobrado uma vez por mês. (Diretoria 2026-06-30.)"""
    out = []
    for md in Medico.objects.filter(eh_preceptor=True).order_by('nome'):
        n = regras.normalizar(md.regra_obs or '')
        if 'mes' in n and 'semana' not in n:
            v = _num_money(md.regra_obs or '')
            if v:
                out.append((md, round(float(v), 2)))
    return out


def _linhas_preceptoria_relatorio(mes):
    """Preceptoria mensal p/ o relatório: uma linha por preceptor no ÚLTIMO DIA do mês
    (mesmo formato de omie.linhas_relatorio_pagar)."""
    ult = _ultimo_dia_mes(mes)
    if not ult:
        return []
    venc = omie.venc_dia10_mes_seguinte(ult)
    cat = omie.CATEGORIA_POR_CLASSE[medplus.CLASSE_PRECEPTORIA]
    return [{'medico': md.nome, 'clinica': '', 'departamento': '',
             'classe': medplus.CLASSE_PRECEPTORIA, 'resumo': 'Preceptoria', 'categoria': cat,
             'valor': valor, 'data': ult.isoformat(), 'vencimento': venc.isoformat()}
            for md, valor in _preceptores_mensais()]


def _linhas_preceptoria_omie(mes):
    """Preceptoria mensal no formato do a pagar OMIE (_escrever) — Fornecedor = CNPJ."""
    ult = _ultimo_dia_mes(mes)
    if not ult:
        return []
    venc = omie.venc_dia10_mes_seguinte(ult)
    cat = omie.CATEGORIA_POR_CLASSE[medplus.CLASSE_PRECEPTORIA]
    return [{'nome': md.cnpj or md.razao_social or md.nome, 'categoria': cat, 'valor': valor,
             'registro': ult, 'vencimento': venc, 'departamento': '',
             'observacao': f'Preceptoria mensal ({mes}) {md.nome}'}
            for md, valor in _preceptores_mensais()]


# Categoria OMIE do ajuste mensal (a diretoria confirma o rótulo exato na OMIE).
_CATEGORIA_AJUSTE = 'Ajuste de Repasse'


def _ajustes_mes(mes):
    """Ajustes (!= 0) do mês, com o médico carregado."""
    return list(AjusteMensal.objects.filter(ano_mes=mes).exclude(valor=0)
                .select_related('medico').order_by('medico__nome'))


def _linhas_ajuste_relatorio(mes):
    """Ajustes do mês p/ o relatório: uma linha por médico no ÚLTIMO DIA do mês."""
    ult = _ultimo_dia_mes(mes)
    if not ult:
        return []
    venc = omie.venc_dia10_mes_seguinte(ult).isoformat()
    return [{'medico': a.medico.nome, 'clinica': '', 'departamento': '', 'classe': 'Ajuste',
             'resumo': 'Ajuste', 'categoria': _CATEGORIA_AJUSTE, 'valor': float(a.valor),
             'data': ult.isoformat(), 'vencimento': venc, 'motivo': a.motivo}
            for a in _ajustes_mes(mes)]


def _linhas_ajuste_omie(mes):
    """Ajustes do mês no formato do a pagar OMIE — Fornecedor = CNPJ. Valor pode ser
    negativo (desconto)."""
    ult = _ultimo_dia_mes(mes)
    if not ult:
        return []
    venc = omie.venc_dia10_mes_seguinte(ult)
    linhas = []
    for a in _ajustes_mes(mes):
        md = a.medico
        obs = f'Ajuste ({mes}) {md.nome}' + (f' — {a.motivo}' if a.motivo else '')
        linhas.append({'nome': md.cnpj or md.razao_social or md.nome, 'categoria': _CATEGORIA_AJUSTE,
                       'valor': float(a.valor), 'registro': ult, 'vencimento': venc,
                       'departamento': '', 'observacao': obs})
    return linhas


def relatorio_mensal(request):
    """Compila os repasses (a pagar) de um MÊS num único xlsx, ordenado por Dr., no
    formato "Repasses em Aberto" (anexo da diretoria). Inclui a preceptoria MENSAL
    (uma linha por preceptor no último dia do mês) e um a pagar OMIE só dela."""
    from datetime import date as _date
    from .services import relatorio
    # Meses disponíveis: derivados das colunas leves de período (sem desserializar o
    # linhas_pagar de TODOS os lotes a cada acesso). (Auditoria 2026-07-02.)
    meses_set = set()
    for ini, fim in (Lote.objects.exclude(periodo_inicio__isnull=True)
                     .values_list('periodo_inicio', 'periodo_fim')):
        fim = fim or ini
        m = _date(ini.year, ini.month, 1)
        while m <= fim:
            meses_set.add(f'{m.year:04d}-{m.month:02d}')
            m = _date(m.year + (m.month == 12), (m.month % 12) + 1, 1)
    meses = sorted(meses_set, reverse=True)
    mes = request.GET.get('mes') or (meses[0] if meses else '')
    linhas_mes = []
    if mes:
        prim, ult = _date(int(mes[:4]), int(mes[5:7]), 1), _ultimo_dia_mes(mes)
        # Só os lotes cujo período toca o mês (as linhas têm data dentro do período);
        # lotes sem período entram por segurança.
        alvo = (Lote.objects.filter(Q(periodo_inicio__lte=ult, periodo_fim__gte=prim)
                                    | Q(periodo_inicio__isnull=True))
                .only('linhas_pagar'))
        for l in alvo:
            linhas_mes.extend(ln for ln in (l.linhas_pagar or [])
                              if (ln.get('data') or '').startswith(mes))
    # Fechamento do mês: preceptoria MENSAL (1 linha por preceptor) + AJUSTES por médico
    # (desconto/acréscimo de meses anteriores) — todos no último dia do mês.
    prec = _linhas_preceptoria_relatorio(mes) if mes else []
    ajustes = _linhas_ajuste_relatorio(mes) if mes else []
    linhas_mes = linhas_mes + prec + ajustes

    baixar = request.GET.get('baixar')
    if baixar == 'omie_fechamento' and (prec or ajustes):
        linhas_omie = _linhas_preceptoria_omie(mes) + _linhas_ajuste_omie(mes)
        res = omie.gerar_contas_pagar_de_linhas(linhas_omie, settings.OMIE_PAGAR_TEMPLATE)
        nome = f'OMIE_Contas_a_Pagar_Fechamento_{relatorio.nome_mes(mes)}.xlsx'
        return FileResponse(io.BytesIO(res.conteudo), as_attachment=True, filename=nome)
    if baixar and linhas_mes:
        titulo = f'Repasses em Aberto - {relatorio.nome_mes(mes)}'
        conteudo = relatorio.gerar_relatorio_mensal(linhas_mes, titulo)
        return FileResponse(io.BytesIO(conteudo), as_attachment=True, filename=f'{titulo}.xlsx')

    resumo, por_medico = Counter(), Counter()
    for ln in linhas_mes:
        v = float(ln.get('valor') or 0)
        resumo[ln.get('resumo') or 'Outros'] += v
        por_medico[ln.get('medico') or '—'] += v
    # Ajustes já salvos por médico (para preencher o formulário de edição).
    ajustes_por_medico = {a.medico_id: a for a in _ajustes_mes(mes)} if mes else {}
    return render(request, 'repasses/relatorio_mensal.html', {
        'meses': [(m, relatorio.nome_mes(m)) for m in meses],
        'mes': mes,
        'nome_mes': relatorio.nome_mes(mes) if mes else '',
        'resumo': [(k, round(v, 2)) for k, v in resumo.most_common()],
        'medicos': [(k, round(v, 2)) for k, v in sorted(por_medico.items())],
        'total': round(sum(resumo.values()), 2),
        'n_linhas': len(linhas_mes),
        'preceptoria_mensal': round(sum(v for _, v in _preceptores_mensais()), 2) if prec else 0,
        'tem_fechamento': bool(prec or ajustes),
        # Formulário de ajustes: todos os médicos + o ajuste salvo (se houver) de cada um.
        'medicos_ajuste': [{'id': m.id, 'nome': m.nome,
                            'valor': ajustes_por_medico[m.id].valor if m.id in ajustes_por_medico else '',
                            'motivo': ajustes_por_medico[m.id].motivo if m.id in ajustes_por_medico else ''}
                           for m in Medico.objects.order_by('nome')],
        'total_ajustes': round(sum(float(a.valor) for a in _ajustes_mes(mes)), 2) if mes else 0,
    })


def salvar_ajuste_mensal(request):
    """Salva os ajustes por médico de um mês (desconto negativo / acréscimo positivo).
    Valor vazio ou 0 apaga o ajuste do médico naquele mês. (Diretoria 2026-07-01.)"""
    if request.method != 'POST':
        raise Http404()
    mes = (request.POST.get('mes') or '').strip()
    if not mes:
        raise Http404('Mês não informado.')
    salvos = 0
    for m in Medico.objects.all():
        valor = _num((request.POST.get(f'ajuste_valor_{m.id}') or '').replace('−', '-'))
        motivo = (request.POST.get(f'ajuste_motivo_{m.id}') or '').strip()
        if valor:   # None ou 0 -> apaga o ajuste do médico
            AjusteMensal.objects.update_or_create(
                ano_mes=mes, medico=m, defaults={'valor': round(valor, 2), 'motivo': motivo})
            salvos += 1
        else:
            AjusteMensal.objects.filter(ano_mes=mes, medico=m).delete()
    messages.success(request, f'{salvos} ajuste(s) salvo(s) para {mes}.')
    return redirect(f"{reverse('repasses:relatorio_mensal')}?mes={mes}")


def lote_detalhe(request, pk):
    """Detalhe de um lote: re-baixar os arquivos, status dos repasses, auditoria."""
    lote = get_object_or_404(Lote, pk=pk)
    db_arqs = list(lote.arquivos.only('grupo', 'nome'))   # nomes p/ os links, sem os bytes
    if db_arqs:
        # Servidos do BANCO — sempre disponíveis (independe da pasta saídas/).
        downloads = [{'grupo': a.grupo, 'arquivo': a.nome, 'existe': True, 'no_banco': True}
                     for a in db_arqs]
        algum_sumiu = False
    else:
        # Lotes antigos (antes do arquivo-no-banco): cai para o disco.
        base = Path(settings.SAIDAS_DIR) / lote.pasta_saida
        downloads = [dict(d, existe=(base / d['arquivo']).is_file(), no_banco=False)
                     for d in lote.downloads]
        algum_sumiu = any(not d['existe'] for d in downloads)
    return render(request, 'repasses/lote_detalhe.html', {
        'lote': lote,
        'downloads': downloads,
        'algum_sumiu': algum_sumiu,
        'tem_no_banco': bool(db_arqs),   # zip "baixar tudo" sai do banco
        'repasses': list(lote.repasses.all()),
        'status_choices': Repasse.STATUS_CHOICES,
    })


def lote_status(request, pk):
    """Salva o andamento (gerado/revisado/enviado/pago) dos repasses do lote."""
    if request.method != 'POST':
        raise Http404()
    lote = get_object_or_404(Lote, pk=pk)
    validos = dict(Repasse.STATUS_CHOICES)
    for r in lote.repasses.all():
        novo = request.POST.get(f'status_{r.id}')
        if novo in validos and novo != r.status:
            r.status = novo
            r.save()
    return redirect('repasses:lote_detalhe', pk=lote.id)


def lote_reabrir(request, pk):
    """Reabre um lote já feito para edição livre na tela de revisão (igual à 1a vez).

    Restaura o arquivo importado e as edições salvas do lote e devolve a revisão;
    ao Exportar de novo, o MESMO lote é atualizado (mesmo token). Bloqueado se o lote
    estiver 'fixado' (algum repasse pago) — a diretoria reverte o status p/ liberar."""
    if request.method != 'POST':
        raise Http404()
    lote = get_object_or_404(Lote, pk=pk)
    if lote.fixado:
        messages.error(request, 'Este lote tem repasse(s) já PAGO(s) e está fixado. '
                       'Para editar, reverta o status do pagamento na lista de repasses abaixo.')
        return redirect('repasses:lote_detalhe', pk=lote.id)
    base = _token_base(lote.token)          # lotes são por DIA; o upload/revisão é um só
    caminho = _caminho_upload(base)
    if caminho is None:
        messages.error(request, 'O relatório importado deste lote não está mais disponível — '
                       'não dá para reabrir. Reimporte o arquivo da MedPlus.')
        return redirect('repasses:lote_detalhe', pk=lote.id)
    # Restaura o rascunho deste lote (a importação de outro arquivo zera os rascunhos).
    RepasseRascunho.objects.update_or_create(
        token=base,
        defaults={'arquivo_nome': lote.arquivo_nome or '', 'dados': lote.edicoes or {}})
    try:
        resultado, aviso, dados = _preparar_revisao(base)
    except medplus.ErroLeituraMedPlus as exc:
        messages.error(request, f'Não foi possível ler o relatório do lote: {exc}')
        return redirect('repasses:lote_detalhe', pk=lote.id)
    info = [f'Editando o lote #{lote.id} ({lote.arquivo_nome or base}). A revisão abre o '
            'arquivo INTEIRO; ao Exportar, todos os dias deste arquivo são atualizados '
            '(um lote por dia).']
    return render(request, 'repasses/revisao.html',
                  _ctx_revisao(resultado, base, aviso, info=info, edicoes=dados))


def lote_excluir(request, pk):
    """Exclui um lote do histórico (e seus repasses/arquivos). Bloqueado se fixado (pago)."""
    if request.method != 'POST':
        raise Http404()
    lote = get_object_or_404(Lote, pk=pk)
    if lote.fixado:
        messages.error(request, 'Este lote tem repasse(s) já PAGO(s) e está fixado — não pode ser '
                       'excluído. Reverta o status do pagamento se precisar mesmo apagá-lo.')
        return redirect('repasses:lote_detalhe', pk=lote.id)
    if lote.periodo_inicio and lote.periodo_inicio == lote.periodo_fim:
        rotulo = f'dia {lote.periodo_inicio:%d/%m/%Y}'      # lote-dia
    else:
        rotulo = lote.arquivo_nome or lote.token
    lote.delete()
    messages.success(request, f'Lote #{pk} ({rotulo}) excluído do histórico. '
                     'Se o dia voltar a ser necessário, reabra outro lote do mesmo '
                     'arquivo e re-exporte, ou reimporte o relatório.')
    return redirect('repasses:lotes')


# --- Correções memorizadas ----------------------------------------------------

def correcoes_lista(request):
    """Lista as correções memorizadas — a memória de ajustes do sistema."""
    itens = list(CorrecaoMemorizada.objects.all())
    classes = list(ClasseMemorizada.objects.all())
    return render(request, 'repasses/correcoes.html', {
        'correcoes': itens,
        'total': len(itens),
        'ativas': sum(1 for c in itens if c.ativo),
        'classes': classes,
        'classes_ativas': sum(1 for c in classes if c.ativo),
    })


def _voltar_correcoes(request):
    """Volta à Administração unificada (aba Correções) se a ação veio de lá; senão à
    página de Correções."""
    if '/administracao' in request.META.get('HTTP_REFERER', ''):
        return redirect(reverse('repasses:administracao') + '?aba=correcoes')
    return redirect('repasses:correcoes')


def correcao_toggle(request, pk):
    """Liga/desliga uma correção (sem apagar — fica o histórico)."""
    if request.method == 'POST':
        c = get_object_or_404(CorrecaoMemorizada, pk=pk)
        c.ativo = not c.ativo
        c.save()
    return _voltar_correcoes(request)


def correcao_remover(request, pk):
    """Remove definitivamente uma correção memorizada."""
    if request.method == 'POST':
        CorrecaoMemorizada.objects.filter(pk=pk).delete()
    return _voltar_correcoes(request)


def classe_toggle(request, pk):
    """Liga/desliga uma classe memorizada (procedimento -> classe)."""
    if request.method == 'POST':
        c = get_object_or_404(ClasseMemorizada, pk=pk)
        c.ativo = not c.ativo
        c.save()
    return _voltar_correcoes(request)


def classe_remover(request, pk):
    """Remove definitivamente uma classe memorizada."""
    if request.method == 'POST':
        ClasseMemorizada.objects.filter(pk=pk).delete()
    return _voltar_correcoes(request)


# --- Regras de repasse (geridas no sistema) -----------------------------------

def administracao(request):
    """Administração unificada: Médicos, Regras e Correções num só lugar (side-nav à
    esquerda, conteúdo numa caixa) e com 'Modo de edição' na mesma tela. (Diretoria
    2026-06-27.)"""
    todos_med = list(Medico.objects.all())
    grupos_med = []
    for codigo, rotulo in Medico.CATEGORIA_CHOICES:
        qs = [m for m in todos_med if m.categoria == codigo]
        if qs:
            grupos_med.append((rotulo, qs))
    todas_reg = list(RegraRepasse.objects.all())
    grupos_reg = []
    for classe, _rotulo in RegraRepasse.CLASSE_CHOICES:
        itens = [r for r in todas_reg if r.classe == classe]
        if itens:
            grupos_reg.append((classe, itens))
    aba = request.GET.get('aba', 'medicos')
    if aba not in ('medicos', 'regras', 'correcoes'):
        aba = 'medicos'
    return render(request, 'repasses/administracao.html', {
        'aba': aba,
        'grupos_medicos': grupos_med,
        'total_medicos': len(todos_med),
        'total_medicos_ativos': sum(1 for m in todos_med if m.ativo),
        'categorias_medico': Medico.CATEGORIA_CHOICES,
        'grupos_regras': grupos_reg,
        'classes_regra': RegraRepasse.CLASSE_CHOICES,
        'correcoes': list(CorrecaoMemorizada.objects.all()),
        'classes': list(ClasseMemorizada.objects.all()),
    })


def regras_lista(request, info=None):
    """Consulta de preços / regras de honorário — listadas e editáveis aqui."""
    todas = list(RegraRepasse.objects.all())
    grupos = []
    for classe, _rotulo in RegraRepasse.CLASSE_CHOICES:
        itens = [r for r in todas if r.classe == classe]
        if itens:
            grupos.append((classe, itens))
    return render(request, 'repasses/regras.html', {
        'grupos': grupos,
        'total': len(todas),
        'ativas': sum(1 for r in todas if r.ativo),
        'info': info,
    })


_CAMPOS_REGRA = ['val_particular', 'val_convenio', 'val_sus', 'val_oci', 'val_cisa']


def regras_salvar(request):
    """Salva as edições de valor feitas direto na página de preços/regras."""
    if request.method != 'POST':
        raise Http404()
    alteradas = 0
    for r in RegraRepasse.objects.all():
        if f'presente_{r.id}' not in request.POST:
            continue  # regra não estava no formulário enviado — não mexe
        mudou = False
        for campo in _CAMPOS_REGRA:
            novo = (request.POST.get(f'{campo}_{r.id}') or '').strip()
            if novo != (getattr(r, campo) or ''):
                setattr(r, campo, novo)
                mudou = True
        ativo = f'ativo_{r.id}' in request.POST
        if ativo != r.ativo:
            r.ativo = ativo
            mudou = True
        if mudou:
            r.save()
            alteradas += 1
    msg = (f'{alteradas} regra(s) atualizada(s) — já valem para os próximos cálculos.'
           if alteradas else 'Nada alterado.')
    if '/administracao' in request.META.get('HTTP_REFERER', ''):
        messages.success(request, msg)
        return redirect(reverse('repasses:administracao') + '?aba=regras')
    return regras_lista(request, info=msg)
