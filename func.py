import pandas as pd
from tkinter import filedialog

def selecionar_arquivo():
    """Função para selecionar um arquivo Excel"""
    arquivo = filedialog.askopenfilename(title="Selecione um arquivo Excel", filetypes=[("Excel Files", "*.xlsx")])
    return arquivo

def organizar_arquivo():
    """Função que organiza um arquivo Excel baseado na escolha do usuário"""
    arquivo = selecionar_arquivo()
    if not arquivo:
        print("Nenhum arquivo foi selecionado.")
        return

    try:
        # Carregar os dados do Excel
        df = pd.read_excel(arquivo)
        print("\nColunas disponíveis:", list(df.columns))

        # Perguntar ao usuário qual coluna usar para organização
        coluna = input("\nDigite o nome da coluna pela qual deseja organizar os dados: ")

        if coluna not in df.columns:
            print("\nColuna inválida. Verifique e tente novamente.")
            return

        # Perguntar se deseja ordem crescente ou decrescente
        ordem = input("\nDeseja organizar em ordem crescente (C) ou decrescente (D)? ").strip().upper()
        crescente = True if ordem == "C" else False

        # Ordenar os dados
        df_organizado = df.sort_values(by=coluna, ascending=crescente)

        # Salvar o novo arquivo Excel
        novo_arquivo = arquivo.replace(".xlsx", "_organizado.xlsx")
        df_organizado.to_excel(novo_arquivo, index=False)

        print(f"\nArquivo organizado e salvo como: {novo_arquivo}")

    except Exception as e:
        print("\nErro ao processar o arquivo:", e)

def criar_nova_planilha():
    """Função que permite o usuário criar uma nova planilha Excel"""
    try:
        # Solicitar as colunas
        colunas = input("\nDigite os nomes das colunas separados por vírgula: ").split(",")
        
        # Criar um dicionário vazio para armazenar os dados
        dados = {col.strip(): [] for col in colunas}

        while True:
            linha = input("\nDigite os valores para cada coluna separados por vírgula (ou 'sair' para finalizar): ")
            if linha.lower() == "sair":
                break
            valores = linha.split(",")
            
            if len(valores) != len(colunas):
                print("\nErro: Número de valores diferente do número de colunas. Tente novamente.")
                continue
            
            for i, col in enumerate(colunas):
                dados[col.strip()].append(valores[i].strip())

        # Criar o DataFrame
        df = pd.DataFrame(dados)

        # Salvar o arquivo
        nome_arquivo = input("\nDigite o nome do arquivo Excel para salvar (ex: tabela.xlsx): ")
        if not nome_arquivo.endswith(".xlsx"):
            nome_arquivo += ".xlsx"

        df.to_excel(nome_arquivo, index=False)

        print(f"\nPlanilha criada e salva como {nome_arquivo}")

    except Exception as e:
        print("\nErro ao criar a planilha:", e)

def somar_colunas():
    """Função para somar valores de colunas"""
    arquivo = selecionar_arquivo()
    if not arquivo:
        print("Nenhum arquivo foi selecionado.")
        return

    try:
        df = pd.read_excel(arquivo)

        # Filtrar apenas colunas numéricas
        colunas_numericas = df.select_dtypes(include=['number']).columns.tolist()
        print("\nColunas numéricas disponíveis:", colunas_numericas)

        if not colunas_numericas:
            print("\nNão há colunas numéricas para somar.")
            return

        # Escolha do usuário
        print("\nOpções para soma:")
        print("1️⃣ Somar todas as colunas numéricas")
        print("2️⃣ Escolher um intervalo de colunas")
        print("3️⃣ Escolher colunas específicas")

        opcao = input("\nDigite o número da opção desejada: ")

        if opcao == "1":
            colunas_selecionadas = colunas_numericas
        elif opcao == "2":
            print("\nColunas disponíveis:", colunas_numericas)
            col_inicio = input("\nDigite o nome da primeira coluna do intervalo: ")
            col_fim = input("Digite o nome da última coluna do intervalo: ")

            if col_inicio in df.columns and col_fim in df.columns:
                idx_inicio = df.columns.get_loc(col_inicio)
                idx_fim = df.columns.get_loc(col_fim)
                if idx_inicio <= idx_fim:
                    colunas_selecionadas = df.columns[idx_inicio:idx_fim+1]
                else:
                    print("\nErro: Ordem das colunas inválida.")
                    return
            else:
                print("\nErro: Uma ou mais colunas não existem no arquivo.")
                return
        elif opcao == "3":
            colunas_especificas = input("\nDigite os nomes das colunas separadas por vírgula: ").split(",")
            colunas_especificas = [col.strip() for col in colunas_especificas]

            if all(col in df.columns for col in colunas_especificas):
                colunas_selecionadas = colunas_especificas
            else:
                print("\nErro: Uma ou mais colunas não existem no arquivo.")
                return
        else:
            print("\nOpção inválida.")
            return

        # Adicionar a coluna "Total" com a soma das colunas selecionadas
        df['Total'] = df[colunas_selecionadas].sum(axis=1)

        print("\n🔢 Resultado da soma por linha:")
        print(df[['Total']])

        # Salvar a versão atualizada do arquivo
        novo_arquivo = arquivo.replace(".xlsx", "_com_total.xlsx")
        df.to_excel(novo_arquivo, index=False)
        print(f"\n📁 Arquivo atualizado salvo como: {novo_arquivo}")

    except Exception as e:
        print("\nErro ao processar o arquivo:", e)
