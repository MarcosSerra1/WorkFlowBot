from model.salvar_dados_em_excel import salvar_dados_no_excel
from model.extrair_dados_planilha import extrair_dados_planilha
from model.converter_para_json import converter_em_arquivo_json
from model.carregar_arquivo_json import carregar_arquivo_json

extracao = extrair_dados_planilha(
    caminho_relativo='WorkFlowBot/Contabilidade_Belem.xlsx')
converter_em_arquivo_json(dados_json=extracao)
salvar_dados_no_excel(
    dados_json=carregar_arquivo_json('WorkFlowBot/dados.json'),
    numero_empresa_desejada=735)
