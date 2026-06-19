import re
from collections import Counter
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from django.conf import settings
from django.http import FileResponse, Http404
from django.shortcuts import render

from .forms import ImportarMedPlusForm
from .models import Medico
from .services import medplus, omie, regras, repasse


def home(request):
    """Tela inicial / painel do sistema de repasses médicos."""
    contexto = {
        'total_medicos': Medico.objects.count(),
        'total_medicos_ativos': Medico.objects.filter(ativo=True).count(),
    }
    return render(request, 'repasses/home.html', contexto)


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
    if base not in caminho.parents or not caminho.is_file():
        return None
    return caminho


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


def _linha_vale(p):
    """Linhas que NÃO entram em nenhum lugar (nem preview): R$0, componentes,
    taxas de sala/utilização e 'Não faturável'."""
    if p.status_calculo in ('nao_recebe', 'componente'):
        return False
    if p.classe == medplus.CLASSE_TAXA:
        return False
    if 'nao faturavel' in regras.normalizar(p.procedimento):
        return False
    return True


def _limpar_linhas(resultado):
    for bloco in resultado.blocos:
        bloco.procedimentos = [p for p in bloco.procedimentos if _linha_vale(p)]
    resultado.blocos = [b for b in resultado.blocos if b.procedimentos]


def _unificar_medicos(resultado):
    """Une blocos do mesmo médico com cadastros MedPlus diferentes — ex.:
    'Dra. Tharcila (PR2)' e 'Dra. Tharcila (Geral)' são a mesma pessoa."""
    canonicos = {}
    novos = []
    for bloco in resultado.blocos:
        nome = _RE_SUFIXO_CADASTRO.sub('', bloco.profissional).strip()
        if nome in canonicos:
            canonicos[nome].procedimentos.extend(bloco.procedimentos)
        else:
            bloco.profissional = nome
            canonicos[nome] = bloco
            novos.append(bloco)
    resultado.blocos = novos


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
        # Residentes não recebem -> não aparecem no preview nem na exportação
        resultado.blocos = [b for b in resultado.blocos
                            if not regras.eh_residente(livro, b.profissional)]
        _aplicar_keiti(resultado)
        _marcar_preceptoria(resultado, livro)
    _limpar_linhas(resultado)
    _indexar(resultado)
    return resultado, aviso


def _separar_por_dia(resultado):
    """Quebra cada bloco (médico) em um bloco por (DIA, CLÍNICA/filial)."""
    novos = []
    for bloco in resultado.blocos:
        grupos = {}
        for p in bloco.procedimentos:
            grupos.setdefault((p.data, p.clinica), []).append(p)

        def _chave(k):
            d, c = k
            return (d or __import__('datetime').date.max, c)

        for (d, c) in sorted(grupos, key=_chave):
            nb = medplus.BlocoMedico(profissional=bloco.profissional)
            nb.data = d
            nb.clinica = c
            nb.procedimentos = grupos[(d, c)]
            novos.append(nb)
    resultado.blocos = novos


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
            if not modo:
                continue  # não preenchido — permanece pendente
            taxa = regras.CATARATA_AVISTA if modo == 'avista' else regras.CATARATA_PARCELADO
            total = taxa * p.valor
            if fellow:
                p.honorario = round(total * (1 - regras.FELLOW_PERCENTUAL), 2)
                p.motivo_calculo = f'Catarata particular ({modo}) — cirurgião 60%; fellow {fellow} 40%.'
                linha = medplus.Procedimento(
                    data=p.data, data_texto=p.data_texto, paciente=p.paciente,
                    procedimento=f'{p.procedimento} (participação em cirurgia)',
                    convenio=p.convenio, quantidade=p.quantidade, valor=p.valor,
                    honorario_medplus=None, classe=medplus.CLASSE_CIRURGIA)
                linha.honorario = round(total * regras.FELLOW_PERCENTUAL, 2)
                linha.status_calculo = 'calculado'
                linha.motivo_calculo = f'Fellow 40% da catarata de {bloco.profissional}.'
                extras[(fellow, p.data)].append(linha)
            else:
                p.honorario = round(total, 2)
                p.motivo_calculo = f'Catarata particular ({modo}) — cirurgião 100% (sem fellow).'
            p.status_calculo = 'calculado'

    for (fellow, data), linhas in extras.items():
        bloco = next((b for b in resultado.blocos
                      if regras.normalizar(b.profissional) == regras.normalizar(fellow)
                      and b.data == data), None)
        if bloco is None:
            m = Medico.objects.filter(nome=fellow).first()
            bloco = medplus.BlocoMedico(profissional=fellow,
                                        razao_social=(m.razao_social if m else ''))
            bloco.data = data
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
        linha.motivo_calculo = 'Repasse de preceptoria semanal.'
        bloco.procedimentos.append(linha)


def _resolver_anestesistas(resultado, post):
    """Para cada cirurgião com cirurgia, registra o anestesista escolhido na
    revisão (valor fixo do dia + horas extras) — gera linha no a pagar e PDF."""
    for i, bloco in enumerate(list(resultado.blocos)):
        if not bloco.tem_cirurgia:
            continue
        nome = (post.get(f'anest_nome_{i}') or '').strip()
        if not nome:
            continue
        horas = _num(post.get(f'anest_horas_{i}')) or 0
        pacientes = _num(post.get(f'anest_pac_{i}'))
        valor = regras.valor_anestesista(nome, horas, pacientes)
        m = Medico.objects.filter(nome=nome).first()
        # Só as CIRURGIAS de fato entram no repasse do anestesista (não os
        # procedimentos de consultório como YAG/laser).
        cirurgias = [p for p in bloco.procedimentos
                     if p.classe == medplus.CLASSE_CIRURGIA and medplus.eh_cirurgia(p.procedimento)]
        datas = [p.data for p in cirurgias if p.data]
        resultado.anestesistas.append({
            'anestesista': nome,
            'razao_social': m.razao_social if m else '',
            'cirurgiao': bloco.profissional,
            'clinica': getattr(bloco, 'clinica', '') or '',
            'data': max(datas) if datas else None,
            'valor': valor,
            'cirurgias': cirurgias,
        })


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
            caminho = _caminho_upload(token)
            try:
                resultado, aviso = _ler_e_processar(caminho, token)
            except medplus.ErroLeituraMedPlus as exc:
                erro = str(exc)
            else:
                return render(request, 'repasses/revisao.html', _ctx_revisao(resultado, token, aviso))
    else:
        form = ImportarMedPlusForm()
    return render(request, 'repasses/importar.html', {'form': form, 'erro': erro})


def _ctx_revisao(resultado, token, aviso, downloads=None, pasta_saida='', pendencias=None):
    cirurgias = [(b.profissional, p) for b in resultado.blocos for p in b.procedimentos
                 if p.classe == medplus.CLASSE_CIRURGIA and p.status_calculo != 'componente']
    return {
        'resultado': resultado,
        'token': token,
        'aviso_regras': aviso,
        'pendencias': pendencias if pendencias is not None else _pendencias_revisao(resultado),
        'downloads': downloads or [],
        'pasta_saida': pasta_saida,
        'qtd_cirurgias': len(cirurgias),
        'classes': medplus.CLASSES,
        'fellows': list(Medico.objects.filter(eh_fellow=True)),
        'anestesistas': list(Medico.objects.filter(eh_anestesista=True)),
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

    resultado, aviso = _ler_e_processar(caminho, token)
    _aplicar_edicoes(resultado, request.POST)
    _resolver_preceptoria(resultado, request.POST)
    _resolver_anestesistas(resultado, request.POST)
    _resolver_cirurgias(resultado, request.POST)

    arquivos, pend = _gerar_arquivos_por_dia(resultado)
    pasta_saida, downloads = _salvar_saidas(arquivos)
    pendencias = _resumir_pendencias(pend)
    ctx = _ctx_revisao(resultado, token, aviso, downloads, pasta_saida, pendencias)
    return render(request, 'repasses/revisao.html', ctx)


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
    receber = omie.gerar_contas_receber(resultado, settings.OMIE_RECEBER_TEMPLATE,
                                        settings.OMIE_CATEGORIA_RECEBER)
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
        arquivos.append((grupo, f'{base}.pdf', repasse.gerar_pdf_anestesista(a, resultado.unidade)))
    return arquivos, pend


def baixar_saida(request, pasta, arquivo):
    """Serve um arquivo gerado da pasta de saídas (com proteção contra path traversal)."""
    base = Path(settings.SAIDAS_DIR).resolve()
    caminho = (base / pasta / arquivo).resolve()
    if base not in caminho.parents or not caminho.is_file():
        raise Http404('Arquivo não encontrado.')
    return FileResponse(open(caminho, 'rb'), as_attachment=True, filename=arquivo)
