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


def _resumir_pendencias(itens):
    """Agrupa pendências repetidas em 'Nx mensagem'."""
    contagem = Counter(itens)
    return [f'{q}× {msg}' if q > 1 else msg for msg, q in contagem.items()]


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


def importar(request):
    """Importa o relatório da MedPlus, classifica, calcula e gera as saídas OMIE."""
    resultado = None
    erro = None
    aviso_regras = None
    downloads = []
    pasta_saida = ''
    pendencias = []

    if request.method == 'POST':
        form = ImportarMedPlusForm(request.POST, request.FILES)
        if form.is_valid():
            arquivo = form.cleaned_data['arquivo']
            try:
                resultado = medplus.ler_relatorio(arquivo, arquivo.name)
            except medplus.ErroLeituraMedPlus as exc:
                erro = str(exc)
            else:
                livro = regras.carregar_livro_padrao()
                if livro is None:
                    aviso_regras = (
                        'A planilha de regras não foi encontrada — os honorários ficaram '
                        '"a definir". Confira o caminho em REGRAS_REPASSE_PATH.'
                    )
                else:
                    regras.processar(resultado, livro)
                    pagar = omie.gerar_contas_pagar(resultado, settings.OMIE_PAGAR_TEMPLATE)
                    receber = omie.gerar_contas_receber(
                        resultado, settings.OMIE_RECEBER_TEMPLATE, settings.OMIE_CATEGORIA_RECEBER)

                    arquivos = [
                        ('Importação OMIE', pagar.nome_arquivo, pagar.conteudo),
                        ('Importação OMIE', receber.nome_arquivo, receber.conteudo),
                    ]
                    for bloco in resultado.blocos:
                        base = repasse.nome_base(bloco)
                        grupo = f'Repasse — {bloco.profissional}'
                        arquivos.append((grupo, f'{base}.xlsx', repasse.gerar_excel(bloco, resultado.unidade)))
                        arquivos.append((grupo, f'{base}.pdf', repasse.gerar_pdf(bloco, resultado.unidade)))

                    pasta_saida, downloads = _salvar_saidas(arquivos)
                    pendencias = _resumir_pendencias(pagar.pendencias + receber.pendencias)
    else:
        form = ImportarMedPlusForm()

    contexto = {
        'form': form,
        'resultado': resultado,
        'erro': erro,
        'aviso_regras': aviso_regras,
        'downloads': downloads,
        'pasta_saida': pasta_saida,
        'pendencias': pendencias,
        'classe_indefinida': medplus.CLASSE_INDEFINIDA,
    }
    return render(request, 'repasses/importar.html', contexto)


def baixar_saida(request, pasta, arquivo):
    """Serve um arquivo gerado da pasta de saídas (com proteção contra path traversal)."""
    base = Path(settings.SAIDAS_DIR).resolve()
    caminho = (base / pasta / arquivo).resolve()
    if base not in caminho.parents or not caminho.is_file():
        raise Http404('Arquivo não encontrado.')
    return FileResponse(open(caminho, 'rb'), as_attachment=True, filename=arquivo)
