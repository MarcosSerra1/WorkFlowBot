import openpyxl
import os

def salvar_dados_no_excel(dados_json : dict, numero_empresa_desejada : int) -> None:

    # Número da empresa desejada
    # numero_empresa_desejada = 763

    # Lista para armazenar os dados dos funcionários da empresa
    dados_funcionarios = []

    # Iterar sobre os dados e encontrar os funcionários da empresa desejada
    for indice, empresa in dados_json["EMPRESA"].items():
        if empresa == numero_empresa_desejada:
            funcionario = {}
            for chave, valor in dados_json.items():
                funcionario[chave] = valor[indice]
            dados_funcionarios.append(funcionario)

    # Obter o diretório atual do script
    diretorio_atual = os.path.dirname(os.path.abspath('.\\'))

    # Caminho planilha Modelo
    caminho_relativo = 'WorkFlowBot/doc/evento_simplificado.xlsx'
    # Construir o caminho absoluto para o arquivo JSON
    caminho_absoluto = os.path.join(diretorio_atual, caminho_relativo)
    # Carregar o arquivo Excel
    planilha = openpyxl.load_workbook(caminho_absoluto)
    # Obtendo a planilha 'evento_simplificado.xls'
    evento_simplificado_sheet = planilha['evento_simplificado.xls']

    # Define a linha inicial onde os dados serão colados na planilha de destino
    linha_destino = 2

    # Itera sobre as linhas da planilha de origem e copia os dados para a planilha de destino
    # for linha in evento_simplificado_sheet.iter_rows(min_row=2):
        # colocando os dados na planilha
    for colaborador in dados_funcionarios:
        # se adiantamento for maior que zero então vamos colocar os dados na planilha
        if colaborador['ADI'] > 0:
            # matrícula
            evento_simplificado_sheet.cell(row=linha_destino, column=1, value=colaborador['MATRÍCULA'])

            # nome
            evento_simplificado_sheet.cell(row=linha_destino, column=2, value=colaborador['COLABORADOR'])

            # tipo de cálculo
            evento_simplificado_sheet.cell(row=linha_destino, column=3, value='Folha')

            # numero do evento ADIANTAMENTO SALARIAL
            evento_simplificado_sheet.cell(row=linha_destino, column=4, value='503')

            # valor manual
            evento_simplificado_sheet.cell(row=linha_destino, column=5, value='Sim')

            # valor do adiantamento
            evento_simplificado_sheet.cell(row=linha_destino, column=6, value=colaborador['ADI'])
        
            # incrementando mais uma linha
            linha_destino += 1
        
        if colaborador['COMISSÕES'] > 0:
            # matrícula
            evento_simplificado_sheet.cell(row=linha_destino, column=1, value=colaborador['MATRÍCULA'])

            # nome
            evento_simplificado_sheet.cell(row=linha_destino, column=2, value=colaborador['COLABORADOR'])

            # tipo de cálculo
            evento_simplificado_sheet.cell(row=linha_destino, column=3, value='Folha')

            # numero do evento COMISSÕES
            evento_simplificado_sheet.cell(row=linha_destino, column=4, value='31')

            # valor manual
            evento_simplificado_sheet.cell(row=linha_destino, column=5, value='Sim')

            # valor da comissão
            evento_simplificado_sheet.cell(row=linha_destino, column=6, value=colaborador['COMISSÕES'])
        
            # incrementando mais uma linha
            linha_destino += 1

        if colaborador['COMISSÕES'] > 0:
            # matrícula
            evento_simplificado_sheet.cell(row=linha_destino, column=1, value=colaborador['MATRÍCULA'])

            # nome
            evento_simplificado_sheet.cell(row=linha_destino, column=2, value=colaborador['COLABORADOR'])

            # tipo de cálculo
            evento_simplificado_sheet.cell(row=linha_destino, column=3, value='Folha')

            # numero do evento DSR
            evento_simplificado_sheet.cell(row=linha_destino, column=4, value='341')

            # valor manual
            evento_simplificado_sheet.cell(row=linha_destino, column=5, value='Não')

            # valor da comissão
            evento_simplificado_sheet.cell(row=linha_destino, column=6, value='0.0')
        
            # incrementando mais uma linha
            linha_destino += 1



    # Salvar o DataFrame em um arquivo Excel
    nome_arquivo_excel = f'evento_simplificado_cmr_{numero_empresa_desejada}.xlsx'
    planilha.save(nome_arquivo_excel)




# Exemplo de uso das funções
# caminho = 'C:/Users/devse/Documentos/GitHub/WorkFlowBot/doc/Contabilidade_Belem.xlsx'
# res = extrair_dados_planilha(caminho=caminho)
# transformar_em_arquivo_json(res)
# print(abrir_arquivo_json('WorkFlowBot/dados.json'))