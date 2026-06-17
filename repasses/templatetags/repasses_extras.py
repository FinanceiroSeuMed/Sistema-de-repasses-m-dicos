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
