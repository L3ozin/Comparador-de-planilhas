import pandas as pd
import os

# Função para formatar o telefone
def formatar_telefone(telefone):
    # Remove qualquer caractere que não seja número
    telefone = ''.join(filter(str.isdigit, str(telefone)))
    
    # Verifica se o telefone tem o comprimento correto (10 ou 11 caracteres)
    if len(telefone) == 10:  # Sem o dígito 9 extra
        return f"({telefone[:2]}) {telefone[2:6]}-{telefone[6:]}"
    elif len(telefone) == 11:  # Com o dígito 9 extra
        return f"({telefone[:2]}) {telefone[2:7]}-{telefone[7:]}"
    return telefone  # Retorna o telefone sem formatação se o comprimento for inválido

def comparar_planilhas(base_path, entrada_path):
    # Carregar a planilha base (formato .xlsx)
    base_df = pd.read_excel(base_path)

    # Carregar a planilha de entrada (formato .csv)
    entrada_df = pd.read_excel(entrada_path)

    # Normalizar os nomes das colunas (remover espaços e caracteres especiais)
    base_df.columns = base_df.columns.str.strip().str.lower()
    entrada_df.columns = entrada_df.columns.str.strip().str.lower()

    # Verificar se as colunas esperadas estão presentes
    colunas_necessarias = {"nome da empresa", "cnpj", "numero de telefone 1", "email", "cidade"}
    if not colunas_necessarias.issubset(
        base_df.columns
    ) or not colunas_necessarias.issubset(entrada_df.columns):
        print("Erro: As colunas esperadas não estão presentes em uma das planilhas.")
        return

    # Formatar os números de telefone em ambas as planilhas antes de qualquer comparação
    base_df["numero de telefone 1"] = base_df["numero de telefone 1"].apply(formatar_telefone)
    entrada_df["numero de telefone 1"] = entrada_df["numero de telefone 1"].apply(formatar_telefone)

    # Mostrar o último valor cadastrado na base antes de qualquer adição
    if not base_df.empty:
        print("Último valor cadastrado na base:")
        print(base_df.iloc[-1])  # Mostra o último registro
    else:
        print("A planilha base está vazia.")

    # Identificar registros únicos na entrada que não estão na base
    novas_informacoes = entrada_df[
        ~entrada_df.apply(tuple, axis=1).isin(base_df.apply(tuple, axis=1))
    ]

    # Adicionar as novas informações à base, se existirem
    if not novas_informacoes.empty:
        print(
            f"{len(novas_informacoes)} novas informações encontradas. Adicionando à base..."
        )
        base_df = pd.concat([base_df, novas_informacoes], ignore_index=True)

        # Salvar a base atualizada no mesmo arquivo Excel
        base_df.to_excel(base_path, index=False, engine="openpyxl")
        print(f"Base atualizada salva em: {base_path}")
    else:
        print("Nenhuma nova informação encontrada. Base permanece inalterada.")

# Caminhos dos arquivos
base_path = "C:\\Users\\leona\\Documents\\Teste myrp\\planilhas\\base.xlsx"  # Caminho da planilha base
entrada_path = "C:\\Users\\leona\\Documents\\Teste myrp\\planilhas\\entrada.xlsx"  # Caminho da planilha de entrada

# Executar o programa
comparar_planilhas(base_path, entrada_path)
