import pandas as pd

arquivo = 'GuiasOCSPSA_411173S-2025.xlsx'
mapa_df = pd.read_excel(arquivo)

resultado = mapa_df.groupby(['CNPJ', 'Plano Interno']).agg({
    'Nome': 'first',
    'Fatura': lambda x: ', '.join(map(str, x.unique())),  # Junta faturas únicas
    'Valor': 'sum'
}).reset_index()

resultado = resultado[['Nome', 'CNPJ', 'Plano Interno', 'Fatura', 'Valor']]

resultado.to_excel("relatorio_por_cnpj.xlsx", index=False)

print("Arquivo 'relatorio_por_cnpj_e_plano.xlsx' criado com sucesso!")
