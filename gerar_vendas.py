import numpy as np
import pandas as pd

# Configuração de semente para dados reproduzíveis
np.random.seed(42)

# Dados base
datas = pd.date_range(start="2026-01-01", end="2026-03-31", freq="D")
produtos = [
    "Notebook",
    "Mouse Sem Fio",
    "Teclado Mecânico",
    "Monitor 27'",
    "Cadeira Ergonomica",
]
regioes = ["Norte", "Sul", "Leste", "Oeste", "Centro-Oeste"]
precos = {
    "Notebook": 3500.00,
    "Mouse Sem Fio": 80.00,
    "Teclado Mecânico": 250.00,
    "Monitor 27'": 1200.00,
    "Cadeira Ergonomica": 900.00,
}

# Geração de 200 registros aleatórios
n_linhas = 200
data_amostra = np.random.choice(datas, size=n_linhas)
prod_amostra = np.random.choice(produtos, size=n_linhas)
regiao_amostra = np.random.choice(regioes, size=n_linhas)
qtd_amostra = np.random.randint(1, 10, size=n_linhas)

# Montagem do DataFrame
df_base = pd.DataFrame(
    {
        "ID_Venda": range(1001, 1001 + n_linhas),
        "Data": data_amostra,
        "Produto": prod_amostra,
        "Regiao": regiao_amostra,
        "Quantidade": qtd_amostra,
        "Preco_Unitario": [precos[p] for p in prod_amostra],
    }
)

# Ordenar vendas por data de forma cronológica
df_base = df_base.sort_values(by="Data").reset_index(drop=True)

# Recalcular IDs de venda em ordem sequencial
df_base["ID_Venda"] = range(1001, 1001 + len(df_base))

# Exportar para Excel com tratamento de erros
try:
    df_base.to_excel("vendas_ficticias.xlsx", index=False)
    print("Planilha 'vendas_ficticias.xlsx' gerada com sucesso!")
except PermissionError:
    print(
        "Erro: Feche a planilha 'vendas_ficticias.xlsx' no Excel e tente novamente."
    )