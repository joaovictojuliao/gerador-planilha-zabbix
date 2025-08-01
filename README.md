# Gerador de Planilha - Zabbix + Flask

Projeto em Python que gera uma planilha `.xlsx` com dados do Zabbix, usando Flask + Gunicorn.

## Como usar
- Roda via `gunicorn run:app -b 0.0.0.0:8080`
- Acesse `http://localhost:8080/gerar-planilha` para baixar a planilha

## Requisitos
- Python 3.10+
- Flask
- pyzabbix
- pandas
- openpyxl
