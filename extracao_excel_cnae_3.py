import pandas as pd
from utilitarios import dt_format_hoje

def extracao_excel_cnae():

    # Carregar o arquivo Excel
    df = pd.read_excel(r"C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_origem\CNAE_Subclasses_2_3_Estrutura_Detalhada.xlsx")

    # Excluir as 5 primeiras linhas
    df = df.iloc[5:].reset_index(drop=True)

    # Filtrar as linhas onde a coluna 4 (correspondente à 'E') não está vazia
    df_filtrado = df[df.iloc[:, 4].notna()]

    # Selecionar as colunas 4 (E) e 5 (F) pelas suas posições
    df_filtrado = df_filtrado.iloc[:, [4, 5]]

    # Renomear as colunas
    df_filtrado.columns = ['CNAE', 'DESCRICAO']

    # Remover qualquer caractere não numérico da coluna 'CNAE'
    df_filtrado['CNAE'] = df_filtrado['CNAE'].astype(str).str.replace(r'\D', '', regex=True)

    # Salvar o DataFrame filtrado em um novo arquivo Excel
    df_filtrado.to_excel(fr'C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_resultantes\3_arquivo_filtrado_{dt_format_hoje}.xlsx', index=False)


extracao_excel_cnae()