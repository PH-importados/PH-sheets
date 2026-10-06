"""
Configuração do gunicorn em produção (Render).

O gunicorn lê este arquivo automaticamente quando é iniciado na raiz do
projeto, então o start command do Render pode continuar `gunicorn app:app`
— os parâmetros ficam versionados aqui, não no painel.
"""
import os

# Render injeta a porta em $PORT (padrão 10000)
bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"

# 2 workers cabem nos 512 MB do plano free (pandas + openpyxl por worker)
workers = 2

# Notas grandes + consulta à API SEFAZ (10 s por chamada) passam dos 30 s padrão
timeout = 120
