import pandas as pd

# Caminho do arquivo Excel
arquivo_excel = "vendas_midias_ml.xlsx"

# Lê a aba original
df = pd.read_excel(arquivo_excel, sheet_name="Original", parse_dates=["DataVenda"])

# Garante que a coluna AnoMes existe
if "AnoMes" not in df.columns:
    df["AnoMes"] = df["DataVenda"].dt.strftime('%Y-%m')

# Calcula a soma da quantidade e do valor dos itens vendidos por mês e por produto
vendas_mensais = (
    df.groupby(["AnoMes", "TipoMidia"])
    .agg(
        QuantidadeVendida=("Quantidade", "sum"),
        ValorTotalItens=("PrecoTotalItem", "sum")
    )
    .reset_index()
    .sort_values(["AnoMes", "TipoMidia"])
)

# Calcula o total de vendas por mês
totais_mes = (
    df.groupby("AnoMes")["ValorTotalPedido"].sum().rename("ValorTotalPedidos_Mes").reset_index()
)

# Adiciona a coluna do valor total do mês para todas as linhas do DataFrame de vendas mensais
vendas_mensais = vendas_mensais.merge(totais_mes, on="AnoMes")

# Salva em uma nova aba
with pd.ExcelWriter(arquivo_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    vendas_mensais.to_excel(writer, index=False, sheet_name="Vendas_Mensais_Produto")
