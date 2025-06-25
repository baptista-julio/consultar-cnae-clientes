# Consulta CNPJ - ReceitaWS

Sistema automatizado para consulta em massa de CNPJs na API [ReceitaWS](https://www.receitaws.com.br/), com validação de CNAEs e geração de relatórios incrementais em Excel. Focado na análise e comparação dos CNAEs dos clientes em relação ao CNAEs da empresa usuário.

<br>

## 📋 Funcionalidades

-   ✅ Consulta automática de CNPJs de clientes via API ReceitaWS
-   ✅ Validação de CNAEs contra lista de CNAEs da empresa
-   ✅ Identificação de porte empresarial (MEI, ME, EPP, DEMAIS)
-   ✅ Processamento incremental com capacidade de pausar e continuar
-   ✅ Geração de comandos SQL para inserção em banco de dados
-   ✅ Relatório em Excel com múltiplas abas (Consultados, Restantes, Erros)
-   ✅ Preservação de zeros à esquerda em CNPJs
-   ✅ Tratamento inteligente de erros e quota da API

<br>

## 🔧 Pré-requisitos

### Softwares Necessários

-   Python 3.8+
-   Oracle Instant Client
-   Acesso ao banco de dados Oracle
-   Conta ativa na ReceitaWS

### Bibliotecas Python

```bash
pip install pandas
pip install requests
pip install oracledb
pip install unidecode
pip install python-dotenv
pip install openpyxl
pip install xlsxwriter
```

<br>

## ⚙️ Configuração

### 1. Oracle Instant Client

-   Baixe o Oracle Instant Client para Windows 64-bit
-   Descompacte em um diretório (ex: `C:\instantclient-basic-windows-64bits-23.7.0.25.01\`)
-   Atualize a variável `caminho_biblioteca_oracle` no código

### 2. Variáveis de Ambiente (.env)

Crie um arquivo `.env` na raiz do projeto:

```env
# Banco de Dados Oracle
USERNAME_ORACLE=seu_usuario
PASSWORD_ORACLE=sua_senha
HOST_ORACLE=host_do_banco
PORT_ORACLE=1521
SERVICE_NAME_ORACLE=nome_do_servico

# ReceitaWS
AUTH_RECEITAWS=Bearer seu_token_aqui
```

### 3. CNAEs da Empresa

Atualize a lista `lista_cnaes_empresa` com os CNAEs que deseja validar:

```python
lista_cnaes_empresa = ["4673700", "4642702", "4649406", ...]
```

<br>

## 🚀 Como Usar

### Execução Básica

```bash
python main.py
```

<br>

### Fluxo de Trabalho

1. **Primeira Execução**

    - Consulta CNPJs no banco de dados Oracle
    - Inicia processamento do zero
    - Cria arquivo Excel: `CNPJ_consulta_incremental_DD-MM-YYYY.xlsx`

2. **Execuções Subsequentes**

    - Detecta arquivo existente automaticamente
    - Continua de onde parou (lê aba "Restantes")
    - Anexa novos dados sem duplicar

3. **Interrupção e Continuação**

    - Pode interromper a qualquer momento (Ctrl+C)
    - Execute novamente para continuar do ponto de parada
    - A numeração sequencial é mantida

<br>

## 📊 Estrutura do Excel

### Aba "Consultados"

| Coluna         | Descrição                                            |
| -------------- | ---------------------------------------------------- |
| CODCLI         | Código do cliente                                    |
| CNPJ           | CNPJ consultado                                      |
| NOME EMPRESA   | Razão social                                         |
| NOME FANTASIA  | Nome fantasia                                        |
| PORTE          | MEI, ME, EPP ou DEMAIS                               |
| SITUACAO CNPJ  | Status na Receita Federal                            |
| CNAE           | Código CNAE                                          |
| DESCRICAO CNAE | Descrição da atividade                               |
| TIPO CNAE      | PRIMARIO ou SECUNDARIO                               |
| IGUALDADE      | IGUAL ou DIFERENTE (comparação com CNAEs da empresa) |
| QTD IGUAL      | Quantidade de CNAEs iguais                           |
| QTD DIFERENTE  | Quantidade de CNAEs diferentes                       |
| COMANDO INSERT | SQL pronto para inserção Oracle                      |

### Aba "Restantes"

-   Lista de CNPJs que ainda precisam ser consultados
-   Atualizada a cada salvamento incremental

### Aba "Erros Consulta"

-   CNPJs que geraram erro durante a consulta
-   Inclui descrição do erro (quota excedida, CNPJ inválido, etc.)

<br>

## 🔄 Processamento Incremental

O sistema salva automaticamente a cada 10 consultas:

-   Anexa novos dados consultados
-   Atualiza lista de restantes
-   Registra novos erros
-   Preserva dados anteriores

<br>

## ⚠️ Tratamento de Erros

### Erros Comuns

1. **"Quota Exceeded"**: Limite da API atingido
2. **"CNPJ inválido"**: CNPJ com formato incorreto
3. **Timeout**: Problema de conexão com a API

### Como o Sistema Trata

-   Registra todos os erros na aba "Erros Consulta"
-   Continua processamento dos demais CNPJs
-   Permite reprocessar erros posteriormente

<br>

## 📝 Consulta SQL Base

```sql
SELECT
    CODCLI,
    REGEXP_REPLACE(cgcent, '[^0-9A-Za-z]', '') AS CNPJ
FROM
    pcclient
WHERE
    CODCLI IN (
        SELECT DISTINCT CODCLI
        FROM pcpedc
        WHERE
            DTFAT >= TRUNC(ADD_MONTHS(SYSDATE, -5), 'MM')
            AND DTFAT <= LAST_DAY(SYSDATE)
            AND CODCLI NOT IN (
                SELECT DISTINCT codcli
                FROM pcfilial
                WHERE codcli IS NOT NULL
            )
    )
    AND TIPOFJ = 'J'
```

<br>

## 🐛 Solução de Problemas

### CNPJs com zero à esquerda cortados

-   O sistema força tratamento como texto no Excel
-   Usa dtype={'CNPJ': str} na leitura

### Arquivo Excel corrompido

-   Delete o arquivo e execute novamente
-   O sistema recriará do zero consultando o banco

<br>

## 📄 Licença

Uso interno da empresa. Não distribuir sem autorização.

<br>

## 💡 Exemplos de Uso

### Saída do Console

```
Arquivo encontrado: CNPJ_consulta_incremental_21-06-2025.xlsx
Carregando 850 CNPJs restantes do arquivo: CNPJ_consulta_incremental_21-06-2025.xlsx
Já foram processados 150 CNPJs
Continuando processamento a partir dos 850 CNPJs restantes

151. CNPJ: 12345678000190
152. CNPJ: 98765432000111
Erro na consulta do CNPJ 98765432000111: Quota Exceeded
153. CNPJ: 11223344000155
...
Arquivo atualizado: CNPJ_consulta_incremental_21-06-2025.xlsx

Resumo da sessão atual:
- Consultados nesta sessão: 8
- Erros nesta sessão: 2
- Total restante: 840
```

<br>

### Comando SQL Gerado

```sql
INSERT INTO tabela (CODCLI, CNPJ, NOME_EMPRESA, NOME_FANTASIA, PORTE, SITUACAO_CNPJ, CNAE, DESCRICAO_CNAE, TIPO_CNAE, IGUALDADE, QTD_IGUAL, QTD_DIFERENTE)
VALUES ('12345', '12345678000190', 'EMPRESA EXEMPLO LTDA', 'EXEMPLO', 'EPP', 'ATIVA', '4673700', 'COMERCIO ATACADISTA DE MATERIAL ELETRICO', 'PRIMARIO', 'IGUAL', '1', '2');
```

<br>

## ❓ FAQ

### P: Como reprocessar apenas os erros?

R: Copie os CNPJs da aba "Erros Consulta" para a aba "Restantes" e execute novamente.

### P: O script funciona com CPFs?

R: Não faria sentido já que as informações são exclusivas aos CNPJs. Então não.

### P: Posso alterar a consulta SQL?

R: Sim, modifique a variável `cnpj_clientes` mantendo as colunas CODCLI e CNPJ.

<br>

## Métricas e Limites

-   **ReceitaWS Free**: 3 consultas/minuto
-   **ReceitaWS Pago**: Varia conforme plano
-   **Salvamento**: A cada 10 registros
-   **Timeout API**: 10 segundos por consulta
-   **Formato CNPJ**: 14 dígitos (zeros à esquerda preservados)

<br>

## Monitoramento

-   Verifique regularmente a aba "Erros Consulta"
-   Monitore consumo da API ReceitaWS
-   Valide integridade dos dados consultados

<br>

## Em caso de dúvidas ou problemas:

-   Verifique os logs de erro no console
-   Consulte a aba "Erros Consulta" no Excel
-   Confirme configurações do `.env`
-   Verifique conectividade com banco Oracle
-   Confirme saldo/quota na ReceitaWS

<br>

## 🎯 Objetivo

Automatizar e facilitar a verificação de similaridade entre as atividades econômicas (CNAEs) dos clientes e da empresa contratante, otimizando o processo de validação cadastral e integração de dados corporativos.

<br>

## 👤 Autor

[Júlio Baptista](https://br.linkedin.com/in/julio-reis-baptista)
Projeto desenvolvido para atualizar os cadastros de clientes da Elétrica Bahiana.
