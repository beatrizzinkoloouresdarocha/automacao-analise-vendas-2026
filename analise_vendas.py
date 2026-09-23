import sys

import matplotlib.pyplot as plt
import pandas as pd

# 1. Carregar Dados com Tratamento de Erros
try:
    df = pd.read_excel("vendas_ficticias.xlsx")
except FileNotFoundError:
    print("\n❌ Erro: O arquivo 'vendas_ficticias.xlsx' não foi encontrado!")
    print("👉 Solução: Execute o script 'gerar_vendas.py' primeiro.\n")
    sys.exit(1)
except pd.errors.EmptyDataError:
    print("\n❌ Erro: O arquivo 'vendas_ficticias.xlsx' está vazio.\n")
    sys.exit(1)
except OSError as e:
    print(f"\n❌ Erro de Leitura/Acesso ao arquivo: {e}\n")
    sys.exit(1)

# 2. Processamento e Novas Colunas
df["Faturamento_Total"] = df["Quantidade"] * df["Preco_Unitario"]
df["Data"] = pd.to_datetime(df["Data"])
df["Mes"] = df["Data"].dt.strftime("%Y-%m")

# 3. Métricas Gerais
faturamento_total = df["Faturamento_Total"].sum()
total_unidades = df["Quantidade"].sum()
ticket_medio = faturamento_total / df["ID_Venda"].nunique()

print("=" * 40)
print(f"Faturamento Total: R$ {faturamento_total:,.2f}")
print(f"Total de Unidades Vendidas: {total_unidades}")
print(f"Ticket Médio por Venda: R$ {ticket_medio:,.2f}")
print("=" * 40)

# 4. Agrupamentos / Tabelas Dinâmicas
vendas_por_produto = (
    df.groupby("Produto")[["Quantidade", "Faturamento_Total"]]
    .sum()
    .sort_values("Faturamento_Total", ascending=False)
)
vendas_por_regiao = (
    df.groupby("Regiao")[["Faturamento_Total"]]
    .sum()
    .sort_values("Faturamento_Total", ascending=False)
)
vendas_por_mes = df.groupby("Mes")[["Faturamento_Total"]].sum()

# 5. Exportação do Relatório Consolidado para Excel
try:
    with pd.ExcelWriter(
        "relatorio_analise_vendas.xlsx", engine="openpyxl"
    ) as writer:
        df.to_excel(writer, sheet_name="Base Tratada", index=False)
        vendas_por_produto.to_excel(writer, sheet_name="Por Produto")
        vendas_por_regiao.to_excel(writer, sheet_name="Por Regiao")
        vendas_por_mes.to_excel(writer, sheet_name="Evolucao Mensal")
except PermissionError:
    print(
        "\n❌ Erro de permissão: O arquivo 'relatorio_analise_vendas.xlsx' está aberto no Excel!"
    )
    print("👉 Por favor, feche o Excel e execute o script novamente.\n")
    sys.exit(1)

# 6. Geração do Gráfico de Faturamento por Produto
plt.figure(figsize=(10, 6))
bars = plt.bar(
    vendas_por_produto.index,
    vendas_por_produto["Faturamento_Total"],
    color="#2b5c8f",
)

# Adicionar rótulos de valor em cima de cada barra
for bar in bars:
    height = bar.get_height()
    plt.text(
        bar.get_x() + bar.get_width() / 2.0,
        height,
        f"R$ {height:,.0f}",
        ha="center",
        va="bottom",
        fontsize=9,
    )

plt.title("Faturamento Total por Produto", fontsize=14, fontweight="bold")
plt.xlabel("Produto", fontsize=11)
plt.ylabel("Faturamento (R$)", fontsize=11)
plt.xticks(rotation=30, ha="right")
plt.tight_layout()

# Salvar o gráfico tratando erro de arquivo bloqueado
try:
    plt.savefig("faturamento_por_produto.png", dpi=300)
except PermissionError:
    print(
        "\n⚠️ Aviso: Não foi possível salvar 'faturamento_por_produto.png' pois a imagem está aberta por outro programa."
    )

plt.show()

print(
    "\n✅ Relatório 'relatorio_analise_vendas.xlsx' e gráfico 'faturamento_por_produto.png' gerados com sucesso!"
)