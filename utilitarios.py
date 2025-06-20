import oracledb
import datetime
from conexoes import DSN
from pandas import DataFrame





# Data Ontem
dt_format_ontem = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime('%d-%m-%Y')

# Data Hoje
dt_format_hoje = (datetime.datetime.today()).strftime('%d-%m-%Y')





# PADRÕES --------------------------------------------------------------------------------------

# Consulta puxando o cabeçalho do OracleDB
def consulta_oracle(query):
    connection = oracledb.connect(DSN) 
    cursor = connection.cursor()
    cursor.execute(
        f"""
            {query}
        """
    )
    resultado = cursor.fetchall()
    cabecalhos = [desc[0] for desc in cursor.description]
    connection.close()
    df_bd = DataFrame(data=resultado, columns=cabecalhos)
    return df_bd


