# Bibliotecas
import requests
import xlsxwriter
from pandas import DataFrame, ExcelWriter
from utilitarios import consulta_oracle, dt_format_hoje
from os import environ
from dotenv import load_dotenv

# Carregando o arquivo .env
load_dotenv()

def extracao_sefaz():
    # Consultando banco de dados e extraindo lista de CNPJs
    df_bd = consulta_oracle("""
                                select 
                                    REGEXP_REPLACE(cgcent, '[^0-9]', '') AS CNPJ
                                from 
                                    pcclient 
                                where 
                                    codcli in (select distinct codcli as cnpj from pcpedc where dTFAT >= '01-jan-2025') 
                                    and codcli <> 1
                                    and tipofj = 'J'

                                """)
    lista_cnpj = df_bd['CNPJ'].tolist()

    # Lista de CNPJs consultados através da consulta SQL
    lista_cnaes_eb = [
        "46.73-7-00", "46.42-7-02", "46.49-4-06",
        "46.51-6-01", "46.72-9-00", "46.79-6-04",
        "46.79-6-99", "47.42-3-00", "47.53-9-00"
    ]

    # DataFrames para armazenar os resultados
    df_inaptas = []
    df_baixadas = []
    df_cnaes_iguais = []
    df_cnaes_diferentes = []

    # Função para consulta CNPJ
    def consulta_receitaws(cnpj):
        days = 7
        url = f"https://receitaws.com.br/v1/cnpj/{cnpj}/days/{days}"
        headers = {
            'Accept': 'application/json',
            'Authorization': environ["AUTH_RECEITAWS"]
        }
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                return response.json()
            else:
                return f"Erro ao consultar CNPJ: {response.status_code}"
        except requests.exceptions.RequestException as e:
            return f"Erro na requisição: {e}"

    count = 1

    # Tratamento dos dados obtidos
    for cnpj in lista_cnpj:
        dados_cnpj = consulta_receitaws(cnpj)
        print(count)

        # Verifica se dados_cnpj é um dicionário
        if isinstance(dados_cnpj, dict):
            nome_empresa = dados_cnpj.get("nome", "Nome não encontrado")
            nome_fantasia = dados_cnpj.get("fantasia", "Nome fantasia não encontrado")
            cnpj_empresa = dados_cnpj.get("cnpj", "CNPJ não encontrado")
            situacao = dados_cnpj.get("situacao", "Situação não encontrada")
            atividades_primarias = dados_cnpj.get("atividade_principal", [{}])[0].get("code", "Atividade principal não encontrada")
            atividades_secundarias = [atividade.get("code", "Atividade não encontrada") for atividade in dados_cnpj.get("atividades_secundarias", [])]

            tudo_do_cnpj = [nome_empresa, nome_fantasia, cnpj_empresa, atividades_primarias, str(atividades_secundarias).replace("[","").replace("]","").replace("'","")]

            if situacao == 'INAPTA':
                df_inaptas.append(tudo_do_cnpj)
            elif situacao == 'BAIXADA':
                df_baixadas.append(tudo_do_cnpj)
            else:
                # Verificar se alguma atividade está na lista CNAEs EB
                atividades = [atividades_primarias] + atividades_secundarias
                cnaes_iguais = any(cnae in lista_cnaes_eb for cnae in atividades)
                if cnaes_iguais:
                    df_cnaes_iguais.append(tudo_do_cnpj)
                else:
                    df_cnaes_diferentes.append(tudo_do_cnpj)

        else:
            # Aqui você pode registrar o erro ou ignorar
            print(f"'Erro ao consultar CNPJ {cnpj}: {dados_cnpj}', a requisição retornou uma string em vez de dicionário.")

        count += 1

    # Criando DataFrames a partir das listas acumuladas
    df_inaptas = DataFrame(df_inaptas, columns=['Nome Empresa', 'Nome Fantasia', 'CNPJ', 'Primária', 'Secundária'])
    df_baixadas = DataFrame(df_baixadas, columns=['Nome Empresa', 'Nome Fantasia', 'CNPJ', 'Primária', 'Secundária'])
    df_cnaes_iguais = DataFrame(df_cnaes_iguais, columns=['Nome Empresa', 'Nome Fantasia', 'CNPJ', 'Primária', 'Secundária'])
    df_cnaes_diferentes = DataFrame(df_cnaes_diferentes, columns=['Nome Empresa', 'Nome Fantasia', 'CNPJ', 'Primária', 'Secundária'])

    # Gravar resultados no arquivo Excel após o loop
    arquivo_excel = fr'C:\eletrica-bahiana\microservicos\correcao-cadastro-clientes-st\arquivos_resultantes\1_empresas_cnae_st_{dt_format_hoje}.xlsx'
    with ExcelWriter(arquivo_excel, engine='xlsxwriter') as writer:
        if not df_inaptas.empty:
            df_inaptas.to_excel(writer, sheet_name='Inaptos', index=False)
        if not df_baixadas.empty:
            df_baixadas.to_excel(writer, sheet_name='Baixadas', index=False)
        if not df_cnaes_iguais.empty:
            df_cnaes_iguais.to_excel(writer, sheet_name='CNAEs Iguais', index=False)
        if not df_cnaes_diferentes.empty:
            df_cnaes_diferentes.to_excel(writer, sheet_name='CNAEs Diferentes', index=False)

    print("Arquivo Excel com 4 abas gerado com sucesso!")

extracao_sefaz()