import pandas as pd
import os 

def extrair_dados_planilha() -> str:
    '''
    Extrai os dados relevantes de uma planilha Excel e os converte em formato JSON.

    Returns:
    - str: Uma string contendo os dados no formato JSON.

    Raises:
    - FileNotFoundError: Se o arquivo especificado não for encontrado.
    - Exception: Se ocorrer algum erro durante a leitura ou conversão dos dados.
    '''

    try:
        # Obter o diretório atual do script
        diretorio_atual = os.getcwd() + os.sep

        # Caminho para a planilha de contabilidade
        caminho_relativo = 'Contabilidade_Belem.xlsx'

        # Construir o caminho absoluto para o arquivo Excel
        caminho_absoluto = os.path.join(diretorio_atual, caminho_relativo)

        # Verificar se o arquivo Excel existe
        if not os.path.isfile(caminho_absoluto):
            raise FileNotFoundError(f'O arquivo "{caminho_absoluto}" não foi encontrado.')

        # Ler o arquivo Excel
        planilha = pd.read_excel(caminho_absoluto)

        # Selecionar as colunas desejadas
        dados_selecionados = planilha[
            ['COLABORADOR', 'COMISSÕES', 'ADI', 'MATRÍCULA', 'EMPRESA']
        ]

        # Converte os dados para JSON
        dados_json = dados_selecionados.to_json()

        return dados_json

    except FileNotFoundError as file_error:
        raise FileNotFoundError(f'O arquivo "{caminho_absoluto}" não foi encontrado.') from file_error
    except Exception as e:
        raise Exception('Ocorreu um erro ao extrair os dados da planilha.') from e
