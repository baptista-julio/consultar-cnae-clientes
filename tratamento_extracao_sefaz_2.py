import pandas as pd
import openpyxl
from utilitarios import consulta_oracle, dt_format_hoje


def tratamento_extracao_sefaz():
    # Carregar o arquivo Excel
    file_path = fr'C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_resultantes\1_empresas_cnae_st_{dt_format_hoje}.xlsx'  # Substitua pelo caminho do seu arquivo
    sheet_name = 'CNAEs Iguais'

    # Dataframe Codcli e CNPJ
    df_codcli_cnpj = consulta_oracle("""
        -- Codcli e CNPJ
        select 
            codcli,
            REGEXP_REPLACE(cgcent, '[^0-9]', '') as CNPJ 
        from 
            pcclient                           
    """) 

    # Ler a aba "CNAEs Iguais"
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Criar uma lista para armazenar os dados formatados
    formatted_data = []

    # Iterar sobre as linhas do DataFrame
    for index, row in df.iterrows():
        nome_empresa = row['Nome Empresa']
        nome_fantasia = row['Nome Fantasia']
        cnpj = str(row['CNPJ']).replace('.', '').replace('/', '').replace('-', '')  # Remover formatação do CNPJ
        primaria = row['Primária']
        secundaria_list = row['Secundária'].split(', ')  # Dividir os CNAEs secundários

        # Adicionar os dados formatados para a CNAE primária
        formatted_data.append({
            'Nome Empresa': nome_empresa,
            'Nome Fantasia': nome_fantasia,
            'CNPJ': cnpj,
            'CNAE': primaria,
            'Tipo': 'Primária'
        })

        # Adicionar os dados formatados para cada CNAE secundário
        for secundaria in secundaria_list:
            formatted_data.append({
                'Nome Empresa': nome_empresa,
                'Nome Fantasia': nome_fantasia,
                'CNPJ': cnpj,
                'CNAE': secundaria,
                'Tipo': 'Secundária'
            })

    # Criar um DataFrame a partir dos dados formatados
    formatted_df = pd.DataFrame(formatted_data)

    # Realizar o merge entre df_codcli_cnpj e formatted_df
    merged_df = pd.merge(formatted_df, df_codcli_cnpj, on='CNPJ', how='left')

    # Reordenar as colunas para que 'codcli' seja a primeira e 'CNPJ' a segunda
    merged_df = merged_df[['CODCLI', 'CNPJ', 'Nome Empresa', 'Nome Fantasia', 'CNAE', 'Tipo']]

    # Lista de CNAEs predefinida
    lista_cnaes_eb = ["46.73-7-00", "46.42-7-02", "46.49-4-06", "46.51-6-01", "46.72-9-00", 
                    "46.79-6-04", "46.79-6-99", "47.42-3-00", "47.53-9-00"]

    # Adicionar a coluna "Igualdade"
    merged_df['Igualdade'] = merged_df['CNAE'].apply(lambda x: 'Igual' if x in lista_cnaes_eb else 'Diferente')

    # Contar quantos CNAEs são "Igual" por CNPJ
    qtd_igualdade = merged_df[merged_df['Igualdade'] == 'Igual'].groupby('CNPJ').size().reset_index(name='Qtd. Igualdade')

    # Realizar o merge para adicionar a contagem ao DataFrame original
    merged_df = pd.merge(merged_df, qtd_igualdade, on='CNPJ', how='left')

    merged_df['CNAE'] = merged_df['CNAE'].astype(str).str.replace(r'\D', '', regex=True)

    # Salvar o resultado em um novo arquivo Excel
    output_file_path = fr'C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_resultantes\2_empresas_cnae_st_formatado_{dt_format_hoje}.xlsx'  # Substitua pelo caminho desejado para o arquivo de saída
    merged_df.to_excel(output_file_path, index=False)

    print(f'Dados formatados e salvos')

tratamento_extracao_sefaz()