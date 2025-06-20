# LIBs
import oracledb
from os import environ
from dotenv import load_dotenv

# Carregando o arquivo .env
load_dotenv()

# Conexão com BD
# client = oracledb.init_oracle_client(lib_dir=r"C:\Users\pedro.quintella\AppData\Local\Programs\Python\Python311\instantclient_21_9")
client = oracledb.init_oracle_client(lib_dir=r"C:\instantclient-basic-windows-64bits-23.7.0.25.01\instantclient_23_7")
username = environ["USERNAME_BD"]
userpwd = environ["PASSWORD_BD"]
host = environ["HOST_BD"]
port = environ["PORT_BD"]
service_name = environ["SERVICE_NAME_BD"]
DSN = environ["USERNAME_BD"]+'/'+environ["PASSWORD_BD"]+'@'+environ["HOST_BD"]+':'+environ["PORT_BD"]+'/'+environ["SERVICE_NAME_BD"]
