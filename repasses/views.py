from django.shortcuts import render

from .forms import ImportarMedPlusForm
from .models import Medico
from .services import medplus, regras


def home(request):
    """Tela inicial / painel do sistema de repasses médicos."""
    contexto = {
        'total_medicos': Medico.objects.count(),
        'total_medicos_ativos': Medico.objects.filter(ativo=True).count(),
    }
    return render(request, 'repasses/home.html', contexto)


def importar(request):
    """Importa o relatório da MedPlus, classifica e exibe os procedimentos."""
    resultado = None
    erro = None
    aviso_regras = None

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
                        'A planilha de regras não foi encontrada — os honorários '
                        'ficaram "a definir". Confira o caminho em REGRAS_REPASSE_PATH.'
                    )
                else:
                    regras.processar(resultado, livro)
    else:
        form = ImportarMedPlusForm()

    contexto = {
        'form': form,
        'resultado': resultado,
        'erro': erro,
        'aviso_regras': aviso_regras,
        'classe_indefinida': medplus.CLASSE_INDEFINIDA,
    }
    return render(request, 'repasses/importar.html', contexto)
