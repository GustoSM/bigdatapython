import pandas as pd

# Caminho do arquivo Excel
arquivo_excel = "vendas_midias_ml.xlsx"

# Abre o arquivo e lista todas as sheets disponíveis
with pd.ExcelFile(arquivo_excel) as xls:
    sheet_names = xls.sheet_names

# Para cada sheet, carrega, altera e salva de volta
with pd.ExcelWriter(arquivo_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    for sheet in sheet_names:
        df = pd.read_excel(arquivo_excel, sheet_name=sheet)
        if "EstadoComprador" in df.columns:
            df["EstadoComprador"] = df["EstadoComprador"].replace("SC", "Santa Catarina")
        df.to_excel(writer, index=False, sheet_name=sheet)
