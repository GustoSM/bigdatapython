import pandas as pd

# Arquivo original
arquivo_csv = "vendas_midias_ml.csv"
arquivo_excel = "vendas_midias_ml.xlsx"

# Lê o arquivo original
df = pd.read_csv(arquivo_csv, sep=";", encoding="utf-8-sig", parse_dates=["DataVenda"])

# Simula junho com dobro de vendas de maio
df_maio = df[df["AnoMes"] == "2025-05"]

df_junho = pd.concat([df_maio, df_maio.copy(deep=True)], ignore_index=True)

# Ajusta as datas
def para_junho(dt):
    return dt.replace(month=6, year=2025)

df_junho["DataVenda"] = pd.to_datetime(df_junho["DataVenda"]).apply(para_junho)
df_junho["AnoMes"] = "2025-06"
df_junho = df_junho.reset_index(drop=True)
df_junho["PedidoID"] = [f"MLJUNHO{100000 + i}" for i in range(len(df_junho))]

# Estoque inicial de cada mídia (maior valor de EstoqueRestante no início de maio)
estoque_inicial_por_tipo = {}
for midia in df_junho["TipoMidia"].unique():
    # O estoque inicial é o maior valor de EstoqueRestante para cada tipo em maio
    estoque_inicial = df_maio[df_maio["TipoMidia"] == midia]["EstoqueRestante"].max()
    estoque_inicial_por_tipo[midia] = estoque_inicial

# Recalcula estoque para junho
df_junho["EstoqueRestante"] = 0
for midia, estoque_inicial in estoque_inicial_por_tipo.items():
    df_midia = df_junho[df_junho["TipoMidia"] == midia]
    idx = df_midia.index
    vendas_cumsum = df_junho.loc[idx, "Quantidade"].cumsum()
    df_junho.loc[idx, "EstoqueRestante"] = estoque_inicial - vendas_cumsum

# Adiciona linhas de estoque final
linhas_estoque_final = []
for midia in df_junho["TipoMidia"].unique():
    estoque_final = df_junho[df_junho["TipoMidia"] == midia]["EstoqueRestante"].iloc[-1]
    linha = {col: "" for col in df_junho.columns}
    linha["TipoMidia"] = midia
    linha["EstoqueRestante"] = estoque_final
    linha["PedidoID"] = "ESTOQUE FINAL"
    linhas_estoque_final.append(linha)

df_junho_com_estoque_final = pd.concat([df_junho, pd.DataFrame(linhas_estoque_final)], ignore_index=True)

# Salva os dois DataFrames em abas separadas
with pd.ExcelWriter(arquivo_excel, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Original")
    df_junho_com_estoque_final.to_excel(writer, index=False, sheet_name="Simulacao_Junho_Dobro")
