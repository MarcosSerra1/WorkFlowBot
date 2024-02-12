
def converter_em_arquivo_json(dados_json : str) -> None:
    '''
    Salva os dados JSON em um arquivo.

    Parameters:
    - dados_json (str): Uma string contendo os dados no formato JSON.

    Raises:
    - Exception: Se ocorrer algum erro ao salvar os dados no arquivo.
    '''

    try:
        # Salvar os dados em um arquivo JSON
        with open('dados.json', 'w', encoding='utf-8') as arquivo:
            arquivo.write(dados_json)

    except Exception as e:
        raise Exception('Ocorreu um erro ao transformar os dados em JSON.') from e
