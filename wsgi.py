import sys
import os

# Garante que o diretório do projeto está no path
sys.path.insert(0, '/var/www/webroot/ROOT')

# Importa a aplicação do run.py
from run import app as application

