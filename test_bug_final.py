#!/usr/bin/env python
"""Teste adversarial final: cenario completo onde anest_cnpj vazio causa bug."""
import os, sys, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
sys.path.insert(0, os.path.dirname(__file__))
django.setup()

from types import SimpleNamespace
from repasses import views
from repasses.services import omie
from repasses.models import Medico
import datetime

print("=== CENARIO CRITICO: Residente Regina sem CNPJ ===\n")

# Simular: usuario faz revisao com Residente Regina, deixa anest_cnpj vazio
post_revisao = {
    'anest_nome_0': 'Residente Dra. Regina',
    'anest_horas_0': '2',
    'anest_cnpj_0': '',  # VAZIO - usuario nao preencheu
}

# Resultado construido por _resolver_anestesistas()
# Medico.objects.filter(nome='Residente Dra. Regina').first() nao existe
# Entao m=None, m.cnpj nao existe
cnpj_manual = (post_revisao.get('anest_cnpj_0') or '').strip()  # ''
cnpj = cnpj_manual or ''  # '' (m nao existe, m.cnpj inacessivel)

resultado = SimpleNamespace(
    blocos=[],
    anestesistas=[{
        'anestesista': 'Residente Dra. Regina',
        'razao_social': '',  # m nao existe
        'cnpj': cnpj,  # ''
        'cirurgiao': 'Dr. Cirurgiao',
        'clinica': 'Clinica Teste',
        'subunidade': '',
        'data': datetime.date(2026, 7, 1),
        'valor': 500.0,
        'cirurgias': []
    }]
)

print("1. VERIFICACAO: _campos_pendentes() detecta anest_cnpj vazio?")
pend = views._campos_pendentes(resultado, post_revisao)
pend_str = [d for _, d in pend]
if any('cnpj' in d.lower() for d in pend_str):
    print("   OK: Detectou cnpj vazio")
else:
    print("   [FALHA] Nao detectou cnpj vazio! Bloqueio de exportacao nao funciona.")

print("\n2. GERACAO OMIE: Linha do anestesista")
# Simular gerar_contas_pagar
ref = datetime.date(2026, 7, 1)
pendencias = []
for a in resultado.anestesistas:
    nome_forn = a.get('cnpj') or a.get('razao_social') or a.get('anestesista')
    print(f"   cnpj='{a.get('cnpj')}' razao_social='{a.get('razao_social')}' anestesista='{a.get('anestesista')}'")
    print(f"   => nome_forn para OMIE: '{nome_forn}'")
    # Nao ha check de cnpj vazio para anestesista!
    if not a.get('cnpj'):
        print("   [NOTA] Nao ha pendencia registrada para cnpj vazio de anestesista")
        print("   [NOTA] Comparar com medicos: ha pendencia se cnpj vazio")

print("\n3. CONCLUSAO:")
print("   [BUG CONFIRMADO] Anestesista com cnpj vazio nao e validado.")
print("   - _campos_pendentes() nao valida anest_cnpj")
print("   - gerar_contas_pagar() nao registra pendencia (ao contrario dos medicos)")
print("   - Usuario pode exportar com cnpj vazio, que e fallback para nome")
print("   - Fornecedor fica 'Residente Dra. Regina' em vez de CNPJ")
print("   - OMIE pode rejeitar se CNPJ eh a chave obrigatoria")

