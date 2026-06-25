import io
import re
from collections import Counter
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from django.conf import settings
from django.contrib import messages
from django.http import FileResponse, Http404
from django.shortcuts import get_object_or_404, redirect, render

from .forms import ImportarMedPlusForm
from .models import (ArquivoSaida, CorrecaoMemorizada, Lote, Medico, RegraRepasse,
                     Repasse, RepasseRascunho)
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
        'total_fellows': sum(1 for m in todos if m.eh_fellow),
        'total_anestesistas': sum(1 for m in todos if m.eh_anestesista),
        'total_preceptores': sum(1 for m in todos if m.eh_preceptor),
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


def _caminho_upload(token: str):
    if not token or not _TOKEN_RE.match(token):
        return None
    base = Path(settings.UPLOADS_DIR).resolve()
    caminho = (base / token).resolve()
    if base not in caminho.parents:
        return None
    if caminho.is_file():
        return caminho
    # Pasta uploads/ limpa? Recupera o .xls de origem guardado no banco (lote).
    lote = Lote.objects.filter(token=token).exclude(upload_conteudo=None).first()
    if lote and lote.upload_conteudo:
        base.mkdir(parents=True, exist_ok=True)
        caminho.write_bytes(bytes(lote.upload_conteudo))
        return caminho
    return None


# --- Rascunho: memória das edições do repasse em andamento ---------------------

# Campos da revisão que ficam salvos (o resto — token, csrf, memorizar — não).
_PREFIXOS_RASCUNHO = ('hon_', 'classe_', 'cat_modo_', 'cat_fellow_',
                      'cat_chefe_', 'cat_fellowval_',
                      'anest_nome_', 'anest_horas_', 'preceptoria_', 'memo_medicos_')


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
        bloco.sel_preceptoria = dados.get(f'preceptoria_{i}', '')
        for p in bloco.procedimentos:
            p.sel_cat_modo = dados.get(f'cat_modo_{p.idx}', '')
            p.sel_cat_fellow = dados.get(f'cat_fellow_{p.idx}', '')
            p.sel_cat_chefe = dados.get(f'cat_chefe_{p.idx}', '')
            p.sel_cat_fellowval = dados.get(f'cat_fellowval_{p.idx}', '')


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
                v = _num(post.get(f'hon_{p.idx}'))
                if v is not None and v >= 0:
                    p.honorario = v


_RE_SUFIXO_CADASTRO = re.compile(r'\s*\([^)]*\)\s*$')
_RE_TITULO_MEDICO = re.compile(r'dra?\b')
# Médicos/agendas deixados de fora por ora (entram em funcionalidade futura)
_EXCLUIR_MEDICOS = ('crivari', 'guilherme', 'maria marta')


def _filtrar_blocos(resultado):
    """Remove agendas que não são de médico (sem Dr./Dra.) e os deixados de fora."""
    novos = []
    for bloco in resultado.blocos:
        n = regras.normalizar(bloco.profissional)
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
            if p.classe == medplus.CLASSE_TAXA:
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
        # Residentes não recebem -> não aparecem no preview nem na exportação
        resultado.blocos = [b for b in resultado.blocos
                            if not regras.eh_residente(livro, b.profissional)]
        _aplicar_keiti(resultado)
        _marcar_preceptoria(resultado, livro)
        _marcar_taxas_sala(resultado)
    _limpar_linhas(resultado)
    _indexar(resultado)
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
    """OCI feito por residente: o residente NÃO recebe; o Dr. Alessander recebe.

    Move as linhas de convênio OCI das agendas de residentes para o repasse do
    Dr. Alessander (agrupando por dia/clínica), recalculando o honorário como se
    ele as tivesse feito. As demais linhas do residente seguem o fluxo normal
    (residente é filtrado depois). Rodar ANTES da filtragem de residentes."""
    from collections import defaultdict
    responsavel = next((m for m in livro.medicos
                        if _OCI_RESPONSAVEL in regras.normalizar(m.nome)), None)
    if responsavel is None:
        return
    extras = defaultdict(list)   # (data, clínica) -> linhas de OCI a transferir
    for bloco in resultado.blocos:
        if not regras.eh_residente(livro, bloco.profissional):
            continue
        mantidas = []
        for p in bloco.procedimentos:
            if regras.mapear_convenio(p.convenio) == 'oci':
                r = regras.calcular(livro, p.procedimento, p.convenio, p.valor,
                                    responsavel.nome, medico_obj=responsavel,
                                    quantidade=p.quantidade)
                p.honorario = r.honorario
                p.status_calculo = r.status
                p.motivo_calculo = (f'OCI feito por residente ({bloco.profissional}) — '
                                    f'repasse do {responsavel.nome}.')
                extras[(bloco.data, bloco.clinica or '')].append(p)
            else:
                mantidas.append(p)
        bloco.procedimentos = mantidas

    alvo = regras.normalizar(responsavel.nome)
    for (data, clinica), linhas in extras.items():
        bloco = next((b for b in resultado.blocos
                      if regras.normalizar(b.profissional) == alvo
                      and b.data == data and (b.clinica or '') == clinica), None)
        if bloco is None:
            bloco = medplus.BlocoMedico(profissional=responsavel.nome,
                                        razao_social=responsavel.razao_social or '')
            bloco.data = data
            bloco.clinica = clinica
            resultado.blocos.append(bloco)
        bloco.procedimentos.extend(linhas)


def _aplicar_keiti(resultado):
    """Dr. Keiti: R$ 1.000 (consultas/exames do dia) + 30% do valor das cirurgias."""
    for bloco in resultado.blocos:
        if 'keiti' not in regras.normalizar(bloco.profissional):
            continue
        tem_exame = False
        novas = []
        for p in bloco.procedimentos:
            if p.classe == medplus.CLASSE_CIRURGIA and p.valor:
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
            if fellow:
                # 60/40 calculado, cada um EDITÁVEL no seu próprio campo:
                #  - o chefe (60%) pelo campo cat_chefe_ na linha do cirurgião;
                #  - o fellow (40%) pelo campo hon_ na PRÓPRIA linha de participação
                #    (aplicado depois, em _aplicar_edicoes_sinteticas). (Diretoria 2026-06-24.)
                chefe_calc = round(total * (1 - regras.FELLOW_PERCENTUAL), 2)
                fellow_calc = round(round(total, 2) - chefe_calc, 2)
                p.chefe_calc = chefe_calc                          # placeholder do cat_chefe
                chefe_ovr = _num(post.get(f'cat_chefe_{p.idx}'))
                manual_chefe = chefe_ovr is not None and chefe_ovr >= 0
                p.honorario = chefe_ovr if manual_chefe else chefe_calc
                p.motivo_calculo = (f'Catarata particular ({modo}) — cirurgião 60%; fellow {fellow} 40%'
                                    + (' (chefe ajustado).' if manual_chefe else '.'))
                linha = medplus.Procedimento(
                    data=p.data, data_texto=p.data_texto, paciente=p.paciente,
                    procedimento=f'{p.procedimento} (participação em cirurgia)',
                    convenio=p.convenio, quantidade=p.quantidade, valor=p.valor,
                    honorario_medplus=None, classe=medplus.CLASSE_CIRURGIA, hora=p.hora)
                linha.honorario = fellow_calc       # 40% padrão (editável na própria linha)
                linha.status_calculo = 'calculado'
                linha.sintetica = True
                linha.editavel = True               # tem campo de honorário próprio
                linha.motivo_calculo = f'Fellow 40% da catarata de {bloco.profissional}.'
                extras[(fellow, p.data, p.clinica)].append(linha)
            else:
                p.honorario = round(total, 2)
                p.motivo_calculo = f'Catarata particular ({modo}) — cirurgião 100% (sem fellow).'
            p.status_calculo = 'calculado'
            p.eh_catarata_part = True   # mantém o seletor à vista/parcelado na tela

    for (fellow, data, clinica), linhas in extras.items():
        bloco = next((b for b in resultado.blocos
                      if regras.normalizar(b.profissional) == regras.normalizar(fellow)
                      and b.data == data and (b.clinica or '') == (clinica or '')), None)
        if bloco is None:
            m = Medico.objects.filter(nome=fellow).first()
            bloco = medplus.BlocoMedico(profissional=fellow,
                                        razao_social=(m.razao_social if m else ''))
            bloco.data = data
            bloco.clinica = clinica
            bloco.participacao = True   # bloco só de participação em catarata (sem caixa de anestesista)
            resultado.blocos.append(bloco)
        bloco.procedimentos.extend(linhas)


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
        resultado.anestesistas.append({
            'indice': i,
            'anestesista': nome,
            'razao_social': m.razao_social if m else '',
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
    for i, bloco in enumerate(resultado.blocos):
        if bloco.tem_cirurgia and not getattr(bloco, 'participacao', False):
            if (post.get(f'anest_nome_{i}') or '').strip() in ('', _VERIFICAR):
                pend.append((f'anest_nome_{i}', f'Anestesista de {bloco.profissional} — confirme (ou "sem anestesista").'))
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
            # Novo repasse importado -> zera a memória de edições do anterior.
            RepasseRascunho.objects.all().delete()
            RepasseRascunho.objects.create(token=token, arquivo_nome=(arquivo.name or ''), dados={})
            try:
                resultado, aviso, dados = _preparar_revisao(token)
            except medplus.ErroLeituraMedPlus as exc:
                erro = str(exc)
            else:
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
        raise Http404('Arquivo da importação não encontrado — refaça o upload.')
    _salvar_rascunho(token, request.POST)
    resultado, aviso, dados = _preparar_revisao(token)
    info = ['✓ Alterações salvas. Pode continuar editando aos poucos — ficam guardadas '
            'até você importar um novo repasse.']
    return render(request, 'repasses/revisao.html',
                  _ctx_revisao(resultado, token, aviso, info=info, edicoes=dados))


def cadastrar_medicos(request):
    """Cadastra os médicos novos (sem cadastro) classificados pelo usuário na revisão.

    O sistema NUNCA assume a categoria: só cadastra quem teve uma categoria escolhida.
    Campos editáveis (razão social, regra) para casos extraordinários. Depois recarrega
    a revisão — já reprocessada com os médicos cadastrados."""
    if request.method != 'POST':
        raise Http404()
    token = request.POST.get('token', '')
    if _caminho_upload(token) is None:
        raise Http404('Arquivo da importação não encontrado — refaça o upload.')
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
        raise Http404('Arquivo da importação não encontrado — refaça o upload.')
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
        raise Http404('Arquivo da importação não encontrado — refaça o upload.')
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
                 info=None, edicoes=None, avisos=None, lote_id=None):
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
        'avisos': avisos or [],
        'edicoes': edicoes or {},
        'downloads': downloads or [],
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
    }


def exportar(request):
    """Passo 2: gera os arquivos a partir do relatório já revisado."""
    if request.method != 'POST':
        raise Http404()
    token = request.POST.get('token', '')
    caminho = _caminho_upload(token)
    if caminho is None:
        raise Http404('Arquivo da importação não encontrado — refaça o upload.')

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
        ctx['info'] = ['⚠️ Confirme os campos destacados (em amarelo) antes de exportar — '
                       'inclusive os "sem anestesista"/"sem fellow", para não passar nada batido.']
        return render(request, 'repasses/revisao.html', ctx)

    info = _memorizar_correcoes(resultado, request.POST)

    arquivos, pend = _gerar_arquivos_por_dia(resultado)
    pasta_saida, downloads = _salvar_saidas(arquivos)
    avisos_dup = _registrar_lote(request, token, resultado, dados, pasta_saida, downloads)
    _guardar_arquivos_no_banco(token, arquivos)        # re-download não depende da pasta saídas/
    _guardar_upload_no_banco(token)                    # re-export sobrevive à limpeza de uploads/
    # Exportado -> larga o cache de edições (vivem no lote agora; reabrir pelo histórico).
    # O "Continuar edição" deixa de oferecer este import. (Diretoria 2026-06-24.)
    RepasseRascunho.objects.filter(token=token).delete()
    pendencias = _resumir_pendencias(pend)
    lote = Lote.objects.filter(token=token).only('id').first()
    ctx = _ctx_revisao(resultado, token, aviso, downloads, pasta_saida, pendencias,
                       info=info, edicoes=dados, avisos=avisos_dup,
                       lote_id=lote.id if lote else None)
    return render(request, 'repasses/revisao.html', ctx)


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

    pagar = omie.gerar_contas_pagar(resultado, settings.OMIE_PAGAR_TEMPLATE)
    receber = omie.gerar_contas_receber(resultado, settings.OMIE_RECEBER_TEMPLATE)
    arquivos.append(('Importação OMIE', 'OMIE_Contas_a_Pagar.xlsx', pagar.conteudo))
    arquivos.append(('Importação OMIE', 'OMIE_Contas_a_Receber.xlsx', receber.conteudo))
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


def _guardar_arquivos_no_banco(token, arquivos):
    """Guarda os bytes dos arquivos gerados no banco, ligados ao lote do token, para
    re-download mesmo se a pasta saídas/ for limpa. Re-export substitui."""
    lote = Lote.objects.filter(token=token).first()
    if not lote:
        return
    lote.arquivos.all().delete()
    ArquivoSaida.objects.bulk_create([
        ArquivoSaida(lote=lote, grupo=g, nome=n, conteudo=conteudo)
        for (g, n, conteudo) in arquivos
    ])


def _guardar_upload_no_banco(token):
    """Guarda o .xls de origem no lote (uma vez), para re-exportar mesmo que a
    pasta uploads/ seja limpa."""
    lote = Lote.objects.filter(token=token).first()
    if not lote or lote.upload_conteudo:
        return
    caminho = _caminho_upload(token)
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
            # agrupa por subpasta (Importação OMIE / Repasses por médico) e evita
            # nomes repetidos no zip.
            pasta = (a.grupo or '').strip().replace('/', '-') or 'Arquivos'
            nome = f'{pasta}/{a.nome}'
            usados[nome] = usados.get(nome, 0) + 1
            if usados[nome] > 1:
                base, _, ext = a.nome.rpartition('.')
                nome = f'{pasta}/{base} ({usados[nome]}).{ext}' if ext else f'{pasta}/{a.nome} ({usados[nome]})'
            zf.writestr(nome, bytes(a.conteudo))
    buf.seek(0)
    rotulo = re.sub(r'[^\w.-]+', '_', (lote.arquivo_nome or f'lote_{pk}')).strip('_') or f'lote_{pk}'
    return FileResponse(buf, as_attachment=True, filename=f'repasses_{rotulo}.zip')


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


def _registrar_lote(request, token, resultado, dados, pasta_saida, downloads):
    """Cria/atualiza o lote (histórico) da exportação e devolve avisos de
    duplicidade (atendimentos que já saíram em lote de outro arquivo)."""
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
    arq_nome = _arquivo_nome(token)
    quem = request.user.get_username() if request.user.is_authenticated else 'diretoria'

    # Duplicidade: atendimentos que já apareceram em lote de OUTRO upload (token).
    # A re-exportação do mesmo upload é excluída pelo exclude(token=token); NÃO
    # pulamos por nome de arquivo (dois uploads podem ter o mesmo nome). Conta por
    # multiconjunto para não subnotificar atendimentos repetidos.
    avisos, novos = [], Counter(fps)
    # só os campos lidos — NÃO arrasta o upload_conteudo (BinaryField pesado).
    outros = Lote.objects.exclude(token=token).only('id', 'arquivo_nome', 'criado_em', 'fingerprints')
    for outro in outros:
        inter = novos & Counter(outro.fingerprints or [])
        n = sum(inter.values())
        if n:
            avisos.append(f'⚠️ {n} atendimento(s) já saíram no lote #{outro.id} '
                          f'({outro.arquivo_nome or "?"}, {outro.criado_em:%d/%m/%Y}) — '
                          'confira para não pagar 2×.')

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
    lote, _ = Lote.objects.update_or_create(token=token, defaults=defaults)
    _sync_repasses(lote, resultado)
    return avisos


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
    for a in resultado.anestesistas:
        Repasse.objects.update_or_create(
            lote=lote, tipo='anestesista', medico=a['anestesista'], data=a.get('data'),
            clinica=a.get('clinica', '') or '',
            defaults={'valor': round(float(a.get('valor') or 0), 2),
                      'razao_social': a.get('razao_social', '') or ''})
        atuais.add(('anestesista', a['anestesista'], a.get('data'), a.get('clinica', '') or ''))
    for r in lote.repasses.all():
        if (r.tipo, r.medico, r.data, r.clinica) not in atuais:
            r.delete()


def lotes_lista(request):
    """Histórico de lotes processados, com o que ainda falta pagar."""
    lotes = list(Lote.objects.all())
    pendentes_total = 0
    for l in lotes:
        reps = list(l.repasses.all())
        l.n_total = len(reps)
        l.n_pagos = sum(1 for r in reps if r.status == Repasse.STATUS_PAGO)
        l.n_pendentes = l.n_total - l.n_pagos
        pendentes_total += l.n_pendentes
    return render(request, 'repasses/lotes.html', {
        'lotes': lotes,
        'total': len(lotes),
        'total_pagar': sum((l.total_pagar for l in lotes), 0),
        'pendentes_total': pendentes_total,
    })


def relatorio_mensal(request):
    """Compila os repasses (a pagar) de um MÊS num único xlsx, ordenado por Dr.,
    no formato "Repasses em Aberto" (anexo da diretoria)."""
    from .services import relatorio
    todas = []
    for l in Lote.objects.only('linhas_pagar'):
        todas.extend(l.linhas_pagar or [])
    meses = sorted({(ln.get('data') or '')[:7] for ln in todas if ln.get('data')}, reverse=True)
    mes = request.GET.get('mes') or (meses[0] if meses else '')
    linhas_mes = [ln for ln in todas if (ln.get('data') or '').startswith(mes)] if mes else []

    if request.GET.get('baixar') and linhas_mes:
        titulo = f'Repasses em Aberto - {relatorio.nome_mes(mes)}'
        conteudo = relatorio.gerar_relatorio_mensal(linhas_mes, titulo)
        return FileResponse(io.BytesIO(conteudo), as_attachment=True, filename=f'{titulo}.xlsx')

    resumo, por_medico = Counter(), Counter()
    for ln in linhas_mes:
        v = float(ln.get('valor') or 0)
        resumo[ln.get('resumo') or 'Outros'] += v
        por_medico[ln.get('medico') or '—'] += v
    return render(request, 'repasses/relatorio_mensal.html', {
        'meses': [(m, relatorio.nome_mes(m)) for m in meses],
        'mes': mes,
        'nome_mes': relatorio.nome_mes(mes) if mes else '',
        'resumo': [(k, round(v, 2)) for k, v in resumo.most_common()],
        'medicos': [(k, round(v, 2)) for k, v in sorted(por_medico.items())],
        'total': round(sum(resumo.values()), 2),
        'n_linhas': len(linhas_mes),
    })


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
    if request.method == 'POST':
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
    caminho = _caminho_upload(lote.token)
    if caminho is None:
        messages.error(request, 'O relatório importado deste lote não está mais disponível — '
                       'não dá para reabrir. Reimporte o arquivo da MedPlus.')
        return redirect('repasses:lote_detalhe', pk=lote.id)
    # Restaura o rascunho deste lote (a importação de outro arquivo zera os rascunhos).
    RepasseRascunho.objects.update_or_create(
        token=lote.token,
        defaults={'arquivo_nome': lote.arquivo_nome or '', 'dados': lote.edicoes or {}})
    try:
        resultado, aviso, dados = _preparar_revisao(lote.token)
    except medplus.ErroLeituraMedPlus as exc:
        messages.error(request, f'Não foi possível ler o relatório do lote: {exc}')
        return redirect('repasses:lote_detalhe', pk=lote.id)
    info = [f'✏️ Editando o lote #{lote.id} ({lote.arquivo_nome or lote.token}). '
            'Faça as alterações e clique em Exportar para atualizar este mesmo lote.']
    return render(request, 'repasses/revisao.html',
                  _ctx_revisao(resultado, lote.token, aviso, info=info, edicoes=dados))


def lote_excluir(request, pk):
    """Exclui um lote do histórico (e seus repasses/arquivos). Bloqueado se fixado (pago)."""
    if request.method != 'POST':
        raise Http404()
    lote = get_object_or_404(Lote, pk=pk)
    if lote.fixado:
        messages.error(request, 'Este lote tem repasse(s) já PAGO(s) e está fixado — não pode ser '
                       'excluído. Reverta o status do pagamento se precisar mesmo apagá-lo.')
        return redirect('repasses:lote_detalhe', pk=lote.id)
    rotulo = lote.arquivo_nome or lote.token
    lote.delete()
    messages.success(request, f'Lote #{pk} ({rotulo}) excluído do histórico.')
    return redirect('repasses:lotes')


# --- Correções memorizadas ----------------------------------------------------

def correcoes_lista(request):
    """Lista as correções memorizadas — a memória de ajustes do sistema."""
    itens = list(CorrecaoMemorizada.objects.all())
    return render(request, 'repasses/correcoes.html', {
        'correcoes': itens,
        'total': len(itens),
        'ativas': sum(1 for c in itens if c.ativo),
    })


def correcao_toggle(request, pk):
    """Liga/desliga uma correção (sem apagar — fica o histórico)."""
    if request.method == 'POST':
        c = get_object_or_404(CorrecaoMemorizada, pk=pk)
        c.ativo = not c.ativo
        c.save()
    return redirect('repasses:correcoes')


def correcao_remover(request, pk):
    """Remove definitivamente uma correção memorizada."""
    if request.method == 'POST':
        CorrecaoMemorizada.objects.filter(pk=pk).delete()
    return redirect('repasses:correcoes')


# --- Regras de repasse (geridas no sistema) -----------------------------------

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
    return regras_lista(request, info=(f'✓ {alteradas} regra(s) atualizada(s) — já valem para os próximos cálculos.'
                                       if alteradas else 'Nada alterado.'))
