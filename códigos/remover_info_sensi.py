import pandas as pd

# Caminho do arquivo Excel existente
arquivo_excel = "vendas_midias_ml.xlsx"

# Lista de colunas sensíveis a serem removidas
colunas_sensiveis = ["PedidoID", "ClienteID", "Marketplace", "TipoVendedor"]

# Lê as duas sheets
df_original = pd.read_excel(arquivo_excel, sheet_name="Original")
df_junho = pd.read_excel(arquivo_excel, sheet_name="Simulacao_Junho_Dobro")

# Remove as colunas sensíveis (se existirem)
df_original = df_original.drop(columns=[col for col in colunas_sensiveis if col in df_original.columns])
df_junho = df_junho.drop(columns=[col for col in colunas_sensiveis if col in df_junho.columns])

# Salva as duas sheets modificadas em um novo arquivo Excel (ou sobrescreve o antigo)
with pd.ExcelWriter(arquivo_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_original.to_excel(writer, index=False, sheet_name="Original")
    df_junho.to_excel(writer, index=False, sheet_name="Simulacao_Junho_Dobro")

print("✔ Dados sensíveis removidos das sheets 'Original' e 'Simulacao_Junho_Dobro'.")