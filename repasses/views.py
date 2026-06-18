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


def _ler_e_processar(caminho, nome=''):
    resultado = medplus.ler_relatorio(str(caminho), nome)
    livro = regras.carregar_livro_padrao()
    aviso = None
    if livro is None:
        aviso = ('A planilha de regras não foi encontrada — os honorários ficaram '
                 '"a definir". Confira REGRAS_REPASSE_PATH.')
    else:
        regras.processar(resultado, livro)
    return resultado, aviso


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
                 if p.classe == medplus.CLASSE_CIRURGIA]
    return {
        'resultado': resultado,
        'token': token,
        'aviso_regras': aviso,
        'pendencias': pendencias if pendencias is not None else _pendencias_revisao(resultado),
        'downloads': downloads or [],
        'pasta_saida': pasta_saida,
        'qtd_cirurgias': len(cirurgias),
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

    pagar = omie.gerar_contas_pagar(resultado, settings.OMIE_PAGAR_TEMPLATE)
    receber = omie.gerar_contas_receber(resultado, settings.OMIE_RECEBER_TEMPLATE,
                                        settings.OMIE_CATEGORIA_RECEBER)
    arquivos = [
        ('Importação OMIE', pagar.nome_arquivo, pagar.conteudo),
        ('Importação OMIE', receber.nome_arquivo, receber.conteudo),
    ]
    for bloco in resultado.blocos:
        if not repasse.pagaveis(bloco):
            continue  # Residentes / sem honorário não geram repasse
        base = repasse.nome_base(bloco)
        grupo = f'Repasse — {bloco.profissional}'
        arquivos.append((grupo, f'{base}.xlsx', repasse.gerar_excel(bloco, resultado.unidade)))
        arquivos.append((grupo, f'{base}.pdf', repasse.gerar_pdf(bloco, resultado.unidade)))

    pasta_saida, downloads = _salvar_saidas(arquivos)
    pendencias = _resumir_pendencias(pagar.pendencias + receber.pendencias)
    ctx = _ctx_revisao(resultado, token, aviso, downloads, pasta_saida, pendencias)
    return render(request, 'repasses/revisao.html', ctx)


def baixar_saida(request, pasta, arquivo):
    """Serve um arquivo gerado da pasta de saídas (com proteção contra path traversal)."""
    base = Path(settings.SAIDAS_DIR).resolve()
    caminho = (base / pasta / arquivo).resolve()
    if base not in caminho.parents or not caminho.is_file():
        raise Http404('Arquivo não encontrado.')
    return FileResponse(open(caminho, 'rb'), as_attachment=True, filename=arquivo)
