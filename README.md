<h1 align="center"> Documentação do Bot WorkFlowBot </h1>
<p align="center">
   <img src="http://img.shields.io/static/v1?label=STATUS&message=EM%20DESENVOLVIMENTO&color=RED&style=for-the-badge" #vitrinedev/>
</p>

## Introdução

WorkFlowBot é uma aplicação Python desenvolvida para extrair dados de uma planilha Excel e salvá-los em um novo arquivo Excel com um formato simplificado. O bot permite que o usuário forneça o número da empresa para extrair os dados específicos relacionados a essa empresa.

## Funcionalidades

O bot WorkFlowBot oferece as seguintes funcionalidades:

1. **Extração de Dados**: Extrai os dados relevantes de uma planilha Excel com base no número da empresa fornecido pelo usuário.
2. **Conversão para JSON**: Converte os dados extraídos em formato JSON.
3. **Salvamento em Excel**: Salva os dados em um novo arquivo Excel com um formato simplificado.

## Dependências

O bot WorkFlowBot depende das seguintes bibliotecas Python:

- pandas
- openpyxl
- tkinter

## Estrutura do Código

O código fonte do bot está organizado da seguinte forma:

- **main.py**: Contém a lógica principal do bot, incluindo a interface gráfica usando a biblioteca tkinter.
- **carregar_arquivo_json.py**: Responsável por carregar os dados de um arquivo JSON.
- **converter_para_json.py**: Converte os dados em formato JSON e os salva em um arquivo.
- **extrair_dados_planilha.py**: Extrai os dados relevantes de uma planilha Excel e os converte em formato JSON.
- **salvar_dados_em_excel.py**: Salva os dados em um arquivo Excel com um formato simplificado.

## Utilização

Para usar o bot WorkFlowBot, execute o arquivo `main.py` e siga as instruções na interface gráfica:

1. Insira o número da empresa no campo designado.
2. Clique no botão "Extrair Dados" para iniciar o processo de extração e salvamento.
3. Um novo arquivo Excel será gerado com os dados da empresa desejada.

## Considerações Finais

O bot WorkFlowBot é uma ferramenta útil para automatizar o processo de extração e manipulação de dados de planilhas Excel. Ele oferece uma interface simples e intuitiva para os usuários interagirem com as funcionalidades fornecidas.

# Autores

| [<img src="https://avatars.githubusercontent.com/u/78652932?v=4" width=115><br><sub>Marcos Serra</sub>](https://github.com/marcosserra1)
