import os
import json


def carregar_arquivo_json() -> dict:
    '''
    Abre um arquivo JSON e carrega seu conteúdo em um dicionário Python.

    Parameters:
    - None

    Returns:
    - dict: Um dicionário contendo os dados do arquivo JSON.

    Raises:
    - FileNotFoundError: Se o arquivo especificado não for encontrado.
    - json.JSONDecodeError: Se ocorrer um erro ao decodificar o conteúdo JSON.
    - UnicodeDecodeError: Se ocorrer um erro de decodificação de caracteres Unicode.
    - Exception: Se ocorrer outro erro durante a leitura ou carregamento do arquivo JSON.
    '''

    try:
        # Obter o diretório atual do script
        diretorio_atual = os.getcwd() + os.sep
        
        # Caminho dados.json
        caminho_relativo = 'dados.json'
        # Construir o caminho absoluto para o arquivo JSON
        caminho_absoluto = os.path.join(diretorio_atual, caminho_relativo)

        # Abrir o arquivo JSON
        with open(caminho_absoluto, 'r', encoding='utf-8') as arquivo:
            # Carregar o JSON
            dados_json = json.load(arquivo)
        return dados_json

    except FileNotFoundError as e:
        raise FileNotFoundError(f'O arquivo "{caminho_relativo}" não foi encontrado.') from e
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError('Erro ao decodificar o conteúdo JSON.', caminho_absoluto, 1) from e
    except UnicodeDecodeError as e:
        raise UnicodeDecodeError('Erro de decodificação de caracteres Unicode.') from e
    except Exception as e:
        raise Exception('Ocorreu um erro ao abrir e carregar o arquivo JSON.') from e
