import pandas as pd
from utilitarios import dt_format_hoje

def merge_final():
        
    # Carregar o arquivo Excel
    df_main_formatado = pd.read_excel(fr"C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_resultantes\2_empresas_cnae_st_formatado_{dt_format_hoje}.xlsx")
    df_excel_cnae = pd.read_excel(fr"C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_resultantes\3_arquivo_filtrado_{dt_format_hoje}.xlsx")

    # Merge
    merged_df = pd.merge(df_main_formatado, df_excel_cnae, on='CNAE', how='left')

    # Reordenar as colunas para que 'codcli' seja a primeira e 'CNPJ' a segunda
    merged_df = merged_df[['CODCLI', 'CNPJ', 'Nome Empresa', 'Nome Fantasia', 'CNAE', 'DESCRICAO', 'Tipo', 'Igualdade', 'Qtd. Igualdade']]

    print(merged_df)

    # Salvar o resultado em um novo arquivo Excel
    output_file_path = fr'C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_resultantes\4_merge_final_{dt_format_hoje}.xlsx'  # Substitua pelo caminho desejado para o arquivo de saída
    merged_df.to_excel(output_file_path, index=False)

merge_final()