from django import template

register = template.Library()


@register.filter
def moeda(valor):
    """Formata um número como moeda brasileira: 1234.5 -> 'R$ 1.234,50'."""
    if valor is None or valor == '':
        return '—'
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return valor
    texto = f'{numero:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
    return f'R$ {texto}'


@register.filter
def valor2(valor):
    """Valor com 2 casas e ponto decimal, para o atributo value de inputs.

    O honorário é guardado em precisão cheia (para a soma fechar), mas o campo
    editável mostra só 2 casas. Se o usuário não mexer, volta esse valor de 2
    casas, e a view preserva a precisão original (não sobrescreve)."""
    if valor is None or valor == '':
        return ''
    try:
        return f'{float(valor):.2f}'
    except (TypeError, ValueError):
        return valor


@register.filter
def valor_exato(valor):
    """Valor em precisão cheia (ponto decimal) para um data-attribute.

    O total ao vivo soma estes valores exatos nas linhas não editadas, para o
    total da tela bater com o total exportado (soma sem arredondar por linha)."""
    if valor is None or valor == '':
        return ''
    try:
        return repr(float(valor))
    except (TypeError, ValueError):
        return ''


@register.filter
def slug_status(valor):
    """Sufixo CSS para o status do cálculo."""
    return {
        'calculado': 'ok',
        'nao_recebe': 'zero',
        'a_definir': 'pendente',
        'componente': 'zero',
        'catarata': 'pendente',
    }.get(valor, 'pendente')


@register.filter
def rotulo_status(valor):
    return {
        'calculado': 'Calculado',
        'nao_recebe': 'Não recebe',
        'a_definir': 'A definir',
        'componente': 'Cirurgia (à parte)',
        'catarata': 'Catarata (definir)',
    }.get(valor, 'A definir')


@register.filter
def badge_repasse(valor):
    """Cor do selo de status do repasse (gerado/revisado/enviado/pago)."""
    return {
        'gerado': 'zero',
        'revisado': 'cirurgia',
        'enviado': 'preceptoria',
        'pago': 'ok',
    }.get(valor, 'zero')


@register.filter
def slug_subclasse(valor):
    """Sufixo CSS para a subclasse do preview (Cirurgias × Procedimentos × ...)."""
    return {
        'Cirurgias': 'cirurgia',
        'Procedimentos': 'procedimento',
        'Exames e Consultas': 'exame',
        'Preceptorias': 'preceptoria',
        'A classificar': 'indefinida',
    }.get(valor, 'indefinida')


@register.filter
def slug_classe(valor):
    """Converte o nome da classe num sufixo de CSS estável."""
    mapa = {
        'Cirurgias e Procedimentos': 'cirurgia',
        'Exames e Consultas': 'exame',
        'Preceptoria': 'preceptoria',
        'Taxas de Sala': 'taxa',
        'A classificar': 'indefinida',
    }
    return mapa.get(valor, 'indefinida')
