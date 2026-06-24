# -*- coding: utf-8 -*-
"""Inicia o Sistema de Repasses Médicos localmente (empacotado com PyInstaller).

Ao executar: prepara o banco (migra e, na 1ª vez, popula médicos/regras e cria um
usuário admin), abre o navegador e sobe o servidor. Feche a janela para parar.
"""
import os
import socket
import sys
import threading
import time
import webbrowser

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')

import django  # noqa: E402
django.setup()
from django.core.management import call_command  # noqa: E402

HOST, PORT = '127.0.0.1', 8000
URL = f'http://{HOST}:{PORT}/'


def preparar_banco():
    """Migra e, se o banco estiver vazio, popula o cadastro inicial + admin."""
    call_command('migrate', interactive=False, verbosity=0)
    from repasses.models import Medico
    if not Medico.objects.exists():
        for cmd in ('seed_medicos', 'seed_regras'):
            try:
                call_command(cmd, verbosity=0)
            except Exception as exc:   # seed é "melhor esforço" — o sistema roda sem
                print(f'  (aviso ao popular {cmd}: {exc})')
    from django.contrib.auth import get_user_model
    User = get_user_model()
    if not User.objects.filter(username='admin').exists():
        User.objects.create_superuser('admin', '', 'admin')


def abrir_navegador():
    """Espera a porta responder e abre o navegador uma vez."""
    for _ in range(80):
        try:
            with socket.create_connection((HOST, PORT), 0.5):
                break
        except OSError:
            time.sleep(0.5)
    try:
        webbrowser.open(URL)
    except Exception:
        pass


def main():
    print('=' * 64)
    print('  Sistema de Repasses Médicos — SeuMed')
    print('=' * 64)
    print('  Preparando o banco de dados...')
    preparar_banco()
    threading.Thread(target=abrir_navegador, daemon=True).start()
    print(f'  Pronto! Acesse no navegador:  {URL}')
    print('  Administração: usuário "admin" / senha "admin".')
    print('  >> Para PARAR o servidor, FECHE esta janela. <<')
    print('=' * 64)
    # use_reloader=False: sem auto-reload (essencial no executável empacotado).
    call_command('runserver', f'{HOST}:{PORT}', use_reloader=False)


if __name__ == '__main__':
    main()
