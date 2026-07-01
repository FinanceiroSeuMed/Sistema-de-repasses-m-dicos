#!/usr/bin/env python
import os, sys, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
sys.path.insert(0, os.path.dirname(__file__))
django.setup()

from types import SimpleNamespace
from repasses.services import omie
import datetime

resultado = SimpleNamespace(
    blocos=[],
    anestesistas=[{
        'anestesista': 'Residente Regina',
        'razao_social': '',  
        'cnpj': '',          # VAZIO
        'cirurgiao': 'Dr. Cirurgiao',
        'clinica': 'Clinica Teste',
        'subunidade': '',
        'data': datetime.date(2026, 7, 1),
        'valor': 500.0,
        'cirurgias': []
    }],
    unidade='Matriz'
)

print("=== Analisando linha OMIE com cnpj vazio ===")
ref = datetime.date(2026, 7, 1)
pend = []
linhas = []
for a in getattr(resultado, 'anestesistas', []):
    dia = a.get('data') or ref
    anest = a.get('anestesista', '')
    cirurgiao = a.get('cirurgiao', '')
    obs = f'Repasse {dia.strftime("%d/%m")} {anest}'
    dep = a.get('clinica', '') or ''
    
    # LINHA 261 DO CODIGO - CALCULO DO 'NOME' DO FORNECEDOR
    nome_forn = a.get('cnpj') or a.get('razao_social') or a.get('anestesista')
    
    print(f"  cnpj vazio: '{a.get('cnpj')}'")
    print(f"  razao_social vazio: '{a.get('razao_social')}'")
    print(f"  anestesista: '{a.get('anestesista')}'")
    print(f"  => nome_forn (para OMIE): '{nome_forn}'")
    
    linhas.append({'nome': nome_forn, 'valor': a.get('valor', 0), 'observacao': obs, 'departamento': dep})

print("\nLinhas para OMIE:")
for l in linhas:
    print(f"  nome='{l['nome']}' | valor={l['valor']}")

if linhas and linhas[0]['nome'] == 'Residente Regina':
    print("\nOK: Fallback para nome do anestesista funcionou")
elif linhas and not linhas[0]['nome']:
    print("\nFALHA: Nome do fornecedor esta vazio!")
else:
    print(f"\nOK: Nome do fornecedor e '{linhas[0]['nome']}'")

