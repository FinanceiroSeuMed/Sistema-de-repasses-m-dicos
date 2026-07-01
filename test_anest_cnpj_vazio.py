#!/usr/bin/env python
"""Teste adversarial: anestesista com CNPJ vazio nas pendencias e OMIE."""
import os, sys, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
sys.path.insert(0, os.path.dirname(__file__))
django.setup()

from types import SimpleNamespace
from repasses.services import omie
from repasses import views
import datetime

# Simular resultado com anestesista que TEM cnpj vazio
resultado = SimpleNamespace(
    blocos=[],
    anestesistas=[{
        'anestesista': 'Residente Regina',
        'razao_social': '',  # vazio
        'cnpj': '',          # VAZIO - problema
        'cirurgiao': 'Dr. Cirurgiao',
        'clinica': 'Clinica Teste',
        'subunidade': '',
        'data': datetime.date(2026, 7, 1),
        'valor': 500.0,
        'cirurgias': []
    }]
)

print("=== TESTE 1: Verificar se _campos_pendentes detecta cnpj vazio ===")
post = {'anest_nome_0': 'Residente Regina', 'anest_horas_0': '2', 'anest_cnpj_0': ''}
pend = views._campos_pendentes(resultado, post)
pend_descricoes = [d for c, d in pend]
print("Pendencias detectadas: " + str(pend_descricoes))
if any('cnpj' in d.lower() for d in pend_descricoes):
    print("[OK] CNPJ vazio FOI detectado como pendencia")
else:
    print("[FALHA] BUG: CNPJ vazio NAO foi detectado como pendencia!")

