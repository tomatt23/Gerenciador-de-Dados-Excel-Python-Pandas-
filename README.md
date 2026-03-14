"""
excel_manager.py

Programa simples para ler, visualizar, adicionar, atualizar e salvar dados em um arquivo Excel (.xlsx)
Usa pandas e openpyxl.

Requisitos:
    pip install pandas openpyxl

O programa:
    - Lê dados de "dados.xlsx" (cria um arquivo de exemplo se não existir)
    - Mostra os dados no terminal de forma organizada
    - Permite visualizar, adicionar, atualizar e salvar os dados
    - Usa funções separadas para cada operação
    - Valida entradas do usuário e evita erros se o arquivo não existir
"""

import os
import sys
import pandas as pd

filepath = "C:/Users/SERVIDORi5M/Documents/dados.xlsx"

def create_sample_file(filepath):
    """
    Cria um arquivo Excel de exemplo com algumas colunas e linhas.
    Chamado automaticamente se o arquivo não existir.
    """
    sample = pd.DataFrame([
        {"ID": 1, "Nome": "Matheus", "Idade": 17, "Email": "matheus.andrade@example.com"},
        {"ID": 2, "Nome": "kayqui", "Idade": 18, "Email": "kayqui.godinho@example.com"},
        {"ID": 3, "Nome": "eloah", "Idade": 16, "Email": "eloah.rodrigues@example.com"},
    ])
    try:
        sample.to_excel(filepath, index=False, engine="openpyxl")
        print(f"Arquivo de exemplo '{filepath}' criado com sucesso.")
    except Exception as e:
        print("Erro ao criar arquivo de exemplo:", e)
        sys.exit(1)

def read_excel_file(filepath):
    """
    Lê a planilha Excel e retorna um DataFrame.
    Se o arquivo não existir, cria um arquivo de exemplo e então lê.
    """
    if not os.path.exists(filepath):
        print(f"Arquivo '{filepath}' não encontrado. Vou criar um arquivo de exemplo.")
        create_sample_file(filepath)

    try:
        df = pd.read_excel(filepath, engine="openpyxl")
        return df
    except Exception as e:
        print("Erro ao ler o arquivo Excel:", e)
        return pd.DataFrame()  # retorna DataFrame vazio em caso de erro

def show_data(df):
    """
    Mostra os dados no terminal de forma organizada.
    Mostramos o índice do DataFrame para facilitar atualizações por índice.
    """
    if df.empty:
        print("\nA planilha está vazia.\n")
        return
    # imprime com índice para referência
    print("\n--- Dados atuais ---")
    print(df.to_string(index=True))
    print("--------------------\n")

def add_row(df):
    """
    Adiciona uma nova linha ao DataFrame pedindo valores para cada coluna.
    Se houver uma coluna 'ID' numérica, auto-incrementamos o ID.
    Retorna o DataFrame atualizado e True se houve mudança.
    """
    if df is None:
        df = pd.DataFrame()

    columns = list(df.columns)
    new_row = {}

    if len(columns) == 0:
        # se não houver colunas, pede ao usuário criar colunas primeiro
        print("A planilha não tem colunas. Informe nomes de colunas separados por vírgula (ex: Nome,Idade,Email):")
        cols_input = input("Colunas: ").strip()
        if not cols_input:
            print("Nenhuma coluna informada. Operação cancelada.")
            return df, False
        columns = [c.strip() for c in cols_input.split(",")]
        for c in columns:
            new_row[c] = input(f"Valor para '{c}': ").strip()
        df = pd.DataFrame([new_row])
        print("Linha adicionada com sucesso.")
        return df, True

    # se existir coluna 'ID' e for numérica, auto-gerar
    if "ID" in columns:
        try:
            max_id = int(df["ID"].max())
            new_id = max_id + 1
        except Exception:
            # se não for possível determinar max, pede ID
            new_id = None

    for col in columns:
        if col == "ID" and 'new_id' in locals() and new_id is not None:
            print(f"Atribuindo ID {new_id} automaticamente.")
            new_row["ID"] = new_id
            continue
        val = input(f"Valor para '{col}' (pressione Enter para vazio): ").strip()
        # tenta converter para número quando apropriado (simples)
        if val.isdigit():
            val = int(val)
        new_row[col] = val

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    print("Linha adicionada com sucesso.")
    return df, True

def update_row(df):
    """
    Atualiza uma linha existente selecionada pelo índice mostrado.
    Retorna o DataFrame atualizado e True se houve mudança.
    """
    if df.empty:
        print("Não há dados para atualizar.")
        return df, False

    show_data(df)
    try:
        idx_input = input("Informe o índice da linha que deseja atualizar: ").strip()
        idx = int(idx_input)
    except ValueError:
        print("Índice inválido. Deve ser um número inteiro.")
        return df, False

    if idx < 0 or idx >= len(df):
        print("Índice fora do intervalo.")
        return df, False

    # Para cada coluna, mostra o valor atual e permite alterar
    for col in df.columns:
        current = df.at[idx, col]
        new_val = input(f"'{col}' atual = '{current}'. Novo valor (Enter para manter): ").strip()
        if new_val == "":
            continue  # mantém o valor atual
        # tenta converter para número se for possível
        if new_val.isdigit():
            new_val = int(new_val)
        df.at[idx, col] = new_val

    print(f"Linha {idx} atualizada com sucesso.")
    return df, True

def save_excel_file(df, filepath):
    """
    Salva o DataFrame no arquivo Excel especificado.
    """
    try:
        df.to_excel(filepath, index=False, engine="openpyxl")
        print(f"Alterações salvas em '{filepath}'.")
        return True
    except Exception as e:
        print("Erro ao salvar o arquivo Excel:", e)
        return False

def main_menu(filepath):
    """
    Loop principal com o menu no terminal.
    """
    df = read_excel_file(filepath)
    modified = False

    while True:
        print("Menu:")
        print("1 - Visualizar todos os dados")
        print("2 - Adicionar novos dados")
        print("3 - Atualizar dados existentes")
        print("4 - Salvar alterações")
        print("5 - Sair")
        choice = input("Escolha uma opção (1-5): ").strip()

        if choice == "1":
            show_data(df)

        elif choice == "2":
            df, changed = add_row(df)
            if changed:
                modified = True

        elif choice == "3":
            df, changed = update_row(df)
            if changed:
                modified = True

        elif choice == "4":
            if save_excel_file(df, filepath):
                modified = False

        elif choice == "5":
            if modified:
                save_choice = input("Há alterações não salvas. Deseja salvar antes de sair? (s/n): ").strip().lower()
                if save_choice == "s":
                    saved = save_excel_file(df, FILEPATH)
                    if saved:
                        print("Saindo. Até logo!")
                        break
                    else:
                        print("Falha ao salvar. Voltando ao menu.")
                elif save_choice == "n":
                    print("Saindo sem salvar. Até logo!")
                    break
                else:
                    print("Opção inválida. Voltando ao menu.")
            else:
                print("Saindo. Até logo!")
                break
        else:
            print("Opção inválida. Digite um número entre 1 e 5.")

if __name__ == "__main__":
    try:
        main_menu(filepath)
    except KeyboardInterrupt:
        print("\nPrograma interrompido pelo usuário. Encerrando.")
