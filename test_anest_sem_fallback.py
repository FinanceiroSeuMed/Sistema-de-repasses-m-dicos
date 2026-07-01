#!/usr/bin/env python
import os, sys, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
sys.path.insert(0, os.path.dirname(__file__))
django.setup()

from types import SimpleNamespace
from repasses.services import omie
from repasses import views
import datetime

# PIOR CASO: Residente Regina, nao tem CNPJ, nao tem razao_social, usuario nao digita anest_cnpj
# Pode isto acontecer se:
# 1. Medico nao existe no cadastro (m = None)
# 2. Usuario deixa campo anest_cnpj vazio na revisao

resultado = SimpleNamespace(
    blocos=[],
    anestesistas=[{
        'anestesista': 'Residente Regina',
        'razao_social': '',  # m.razao_social == '' se m nao existe
        'cnpj': '',          # cnpj_manual vazio + m.cnpj vazio = ''
        'cirurgiao': 'Dr. Cirurgiao',
        'clinica': 'Clinica Teste',
        'subunidade': '',
        'data': datetime.date(2026, 7, 1),
        'valor': 500.0,
        'cirurgias': []
    }]
)

print("=== PIOR CASO: Sem CNPJ, Sem razao_social, Sem anestesista? ===")
# Simular caso onde anestesista name field estah vazio
resultado_mal = SimpleNamespace(
    blocos=[],
    anestesistas=[{
        'anestesista': '',  # VAZIO
        'razao_social': '',
        'cnpj': '',
        'cirurgiao': 'Dr. Cirurgiao',
        'clinica': 'Clinica Teste',
        'subunidade': '',
        'data': datetime.date(2026, 7, 1),
        'valor': 500.0,
        'cirurgias': []
    }]
)

print("Teste: anestesista='', razao_social='', cnpj=''")
a = resultado_mal.anestesistas[0]
nome_forn = a.get('cnpj') or a.get('razao_social') or a.get('anestesista')
print(f"  nome_forn = '{nome_forn}'")
if nome_forn == '':
    print("  FALHA: Nome do fornecedor estah vazio! OMIE vai rejeitar!")
else:
    print(f"  OK: nome_forn = '{nome_forn}'")

