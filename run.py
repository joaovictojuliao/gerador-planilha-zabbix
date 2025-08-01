from flask import Flask, send_file, render_template_string
import gerar_planilha

app = Flask(__name__)

@app.route('/')
def home():
    html = """
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <title>Gerador de Planilha Load</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background-color: #f5f5f5;
                text-align: center;
                padding-top: 100px;
            }
            .container {
                background-color: #fff;
                padding: 40px;
                border-radius: 10px;
                box-shadow: 0 0 15px rgba(0,0,0,0.1);
                display: inline-block;
            }
            a.button {
                display: inline-block;
                margin-top: 20px;
                padding: 10px 20px;
                background-color: #007bff;
                color: white;
                text-decoration: none;
                border-radius: 5px;
                font-size: 16px;
            }
            a.button:hover {
                background-color: #0056b3;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h2>🚀 Planilha de Load Saveincloud 🚀</h2>
            <p>Clique no botão abaixo para gerar a planilha:</p>
            <a href="/gerar-planilha" class="button">📊 Gerar Relatório</a>
        </div>
    </body>
    </html>
    """
    return render_template_string(html)

@app.route('/gerar-planilha')
def gerar():
    caminho = gerar_planilha.gerar_excel()  # essa função precisa retornar o caminho completo do arquivo
    return send_file(caminho, as_attachment=True)

if __name__ == '__main__':
    app.run()


