import pandas as pd
import os 


def extrair_dados_planilha(caminho_relativo : str) -> str:
    '''
    Extrai os dados relevantes de uma planilha Excel e os converte em formato JSON.

    Parameters:
    - caminho_relativo (str): O caminho do arquivo Excel a ser lido.

    Returns:
    - str: Uma string contendo os dados no formato JSON.

    Raises:
    - FileNotFoundError: Se o arquivo especificado não for encontrado.
    - Exception: Se ocorrer algum erro durante a leitura ou conversão dos dados.
    '''

    try:
        # Obter o diretório atual do script
        diretorio_atual = os.path.dirname(os.path.abspath('.\\'))

        # Construir o caminho absoluto para o arquivo JSON
        caminho_absoluto = os.path.join(diretorio_atual, caminho_relativo)


        # Ler o arquivo Excel
        planilha = pd.read_excel(caminho_absoluto)

        # Seleciona as colunas desejadas
        dados_selecionados = planilha[
            ['COLABORADOR', 'COMISSÕES', 'ADI', 'MATRÍCULA', 'EMPRESA']
        ]

        # Converte os dados para JSON
        dados_json = dados_selecionados.to_json()

        return dados_json

    except FileNotFoundError as e:
        raise FileNotFoundError(f'O arquivo "{caminho_absoluto}" não foi encontrado.') from e
    except Exception as e:
        raise Exception('Ocorreu um erro ao extrair os dados da planilha.') from e
