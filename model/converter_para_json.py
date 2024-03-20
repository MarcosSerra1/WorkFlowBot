def converter_em_arquivo_json(dados_json: str) -> None:
    '''
    Salva os dados JSON em um arquivo.

    Parameters:
    - dados_json (str): Uma string contendo os dados no formato JSON.

    Raises:
    - IOError: Se ocorrer um erro de E/S durante a operação de gravação do arquivo.
    - Exception: Se ocorrer qualquer outro erro durante a operação.
    '''

    try:
        # Abre o arquivo 'dados.json' para escrita
        with open('dados.json', 'w', encoding='utf-8') as arquivo:
            # Escreve os dados JSON no arquivo
            arquivo.write(dados_json)

    except IOError as io_error:
        # Se houver um erro de E/S (Input/Output), lança a exceção IOError
        raise IOError('Erro de E/S ao salvar os dados no arquivo.') from io_error
    except Exception as e:
        # Se houver qualquer outro tipo de exceção, lança uma exceção genérica
        raise Exception('Ocorreu um erro ao salvar os dados em JSON.') from e
