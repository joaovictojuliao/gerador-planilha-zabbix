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
                background: linear-gradient(to bottom right, #5c6bc0, #42a5f5);
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                color: #fff;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
            }

            .card {
                background-color: #ffffff;
                color: #333;
                border-radius: 12px;
                padding: 40px;
                text-align: center;
                box-shadow: 0 4px 12px rgba(0,0,0,0.3);
                width: 400px;
            }

            .card h2 {
                margin-top: 0;
                font-size: 24px;
            }

            .button {
                display: inline-block;
                margin-top: 20px;
                padding: 12px 24px;
                background-color: #3f51b5;
                color: white;
                text-decoration: none;
                border: none;
                border-radius: 8px;
                font-size: 16px;
                cursor: pointer;
            }

            .button:hover {
                background-color: #303f9f;
            }
        </style>
    </head>
    <body>
        <div class="card">
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
