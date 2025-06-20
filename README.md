# Consulta de CNPJ/CNAE – ReceitaWS + Oracle

Projeto em **Python** para automatizar a consulta de dados de CNPJ pela [ReceitaWS](https://www.receitaws.com.br/), focado na análise e comparação dos CNAEs dos clientes em relação à empresa contratante. O resultado da consulta é exportado em arquivo **Excel**, já contendo os comandos **INSERT** prontos para serem executados no banco de dados **Oracle**.

---

## ⚙️ Funcionalidades

-   Consulta automática de dados cadastrais de empresas (CNPJ) pela ReceitaWS.
-   Extração dos CNAEs principais e secundários dos clientes.
-   Comparação dos CNAEs dos clientes com o da empresa contratante, facilitando a análise de similaridade e adequação.
-   Exportação dos resultados em arquivo Excel, incluindo comandos INSERT prontos para uso no Oracle.
-   Estrutura de scripts simples, ideal para validação cadastral, compliance e integração de dados.

---

## 🛠 Tecnologias Utilizadas

-   **Linguagem:** Python
-   **API:** ReceitaWS
-   **Banco de Dados:** Oracle
-   **Bibliotecas:**
    -   [pandas](https://pandas.pydata.org/)
    -   [requests](https://requests.readthedocs.io/en/latest/)
    -   [os](https://docs.python.org/3/library/os.html)
    -   [python-dotenv](https://saurabh-kumar.com/python-dotenv/)

<!-- ---

## 🚀 Como Usar

1. **Clone este repositório:**

    ```bash
    git clone https://github.com/seu-usuario/seu-repositorio.git
    cd seu-repositorio
    ```

2. **Crie um arquivo `.env`** com as variáveis necessárias (exemplo de chave da API):

    ```
    RECEITAWS_TOKEN=sua_chave_api
    ```

3. **Instale as dependências do projeto:**

    ```bash
    pip install -r requirements.txt
    ```

4. **Adicione a lista de CNPJs a serem consultados** (pode ser em um arquivo `.txt`, `.csv` ou no próprio script, conforme sua implementação).

5. **Execute o script principal:**

    ```bash
    python consulta_cnpj.py
    ```

6. **Confira o arquivo Excel gerado** na pasta do projeto. Ele irá conter:
    - Dados detalhados dos CNPJs consultados.
    - Os comandos INSERT prontos para serem executados no Oracle. -->

---

## 🎯 Objetivo

Automatizar e facilitar a verificação de similaridade entre as atividades econômicas (CNAEs) dos clientes e da empresa contratante, otimizando o processo de validação cadastral e integração de dados corporativos.

---

## 👤 Autor

[Júlio Baptista](https://br.linkedin.com/in/julio-reis-baptista)
Projeto desenvolvido para atualizar os cadastros de clientes da Elétrica Bahiana.
