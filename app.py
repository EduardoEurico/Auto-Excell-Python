from flask import Flask, request, render_template, jsonify
import pandas as pd
import os

app = Flask(__name__, static_folder="static", template_folder="templates")
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/criar_planilha', methods=['GET', 'POST'])
def criar_nova_planilha():
    if request.method == 'POST':
        try:
            colunas = request.form['colunas'].split(",")  # Recebe as colunas do formulário
            dados = {col.strip(): [] for col in colunas}

            linhas = request.form['linhas'].split("\n")  # Cada linha é um conjunto de valores
            
            for linha in linhas:
                valores = linha.split(",")
                if len(valores) != len(colunas):
                    return "Erro: Número de valores diferente do número de colunas. Tente novamente.", 400

                for i, col in enumerate(colunas):
                    dados[col.strip()].append(valores[i].strip())

            df = pd.DataFrame(dados)

            nome_arquivo = request.form['arquivo']
            if not nome_arquivo.endswith(".xlsx"):
                nome_arquivo += ".xlsx"

            df.to_excel(nome_arquivo, index=False)

            return f"Planilha criada e salva como {nome_arquivo} com sucesso!"

        except Exception as e:
            return f"Erro ao criar a planilha: {str(e)}", 500

    return render_template("criarP.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Nenhum arquivo selecionado."}), 400

    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    df = pd.read_excel(file_path)
    return jsonify({"message": "Arquivo enviado com sucesso!", "columns": list(df.columns), "filename": file.filename})

@app.route("/organizar", methods=["POST"])
def organizar_arquivo():
    data = request.json
    file_path = os.path.join(UPLOAD_FOLDER, data["filename"])
    
    if not os.path.exists(file_path):
        return jsonify({"error": "Arquivo não encontrado."}), 400
    
    df = pd.read_excel(file_path)
    coluna = data["column"]
    crescente = data["order"] == "C"

    if coluna not in df.columns:
        return jsonify({"error": "Coluna inválida."}), 400

    df_organizado = df.sort_values(by=coluna, ascending=crescente)
    new_file = file_path.replace(".xlsx", "_organizado.xlsx")
    df_organizado.to_excel(new_file, index=False)

    return jsonify({"message": f"Arquivo organizado salvo como {new_file}!"})

@app.route("/somar_colunas", methods=["POST"])
def somar_colunas():
    data = request.json
    file_path = os.path.join(UPLOAD_FOLDER, data["filename"])
    
    if not os.path.exists(file_path):
        return jsonify({"error": "Arquivo não encontrado."}), 400
    
    df = pd.read_excel(file_path)
    colunas = data["columns"]

    if not all(col in df.columns for col in colunas):
        return jsonify({"error": "Uma ou mais colunas não existem."}), 400

    df["Total"] = df[colunas].sum(axis=1)
    new_file = file_path.replace(".xlsx", "_com_total.xlsx")
    df.to_excel(new_file, index=False)

    return jsonify({"message": f"Arquivo atualizado salvo como {new_file}!"})

if __name__ == "__main__":
    app.run(debug=True)
