import customtkinter as ctk
from tkinter import messagebox
from model.salvar_dados_em_excel import salvar_dados_no_excel
from model.extrair_dados_planilha import extrair_dados_planilha
from model.converter_para_json import converter_em_arquivo_json
from model.carregar_arquivo_json import carregar_arquivo_json


def extrair_dados() -> None:
    '''
    Extrai dados de uma planilha com base no número da empresa fornecido pelo usuário.

    Parameters:
    None
    
    Returns:
    None
    
    Raises:
    None

    Esta função extrai dados de uma planilha com base no número da empresa fornecido pelo usuário.
    Primeiro, recupera o número da empresa digitado pelo usuário no campo de entrada 'entry'.
    Verifica se o número da empresa é um número inteiro válido.
    Se não for um número inteiro válido, exibe uma mensagem de erro e encerra a função.
    Caso contrário, converte o número da empresa em um inteiro.
    Em seguida, extrai os dados da planilha usando a função extrair_dados_planilha().
    Converte os dados extraídos em formato JSON usando a função converter_em_arquivo_json().
    Carrega os dados JSON convertidos do arquivo usando a função carregar_arquivo_json().
    Finalmente, salva os dados no Excel com base no número da empresa desejada.
    '''
    try:
        empresa = entry.get()
        if not empresa.isdigit():
            raise ValueError("Insira um número inteiro válido.")

        empresa = int(empresa)
        extracao = extrair_dados_planilha()
        converter_em_arquivo_json(dados_json=extracao)
        salvar_dados_no_excel(
            dados_json=carregar_arquivo_json(),
            numero_empresa_desejada=empresa)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")



# Configuração da janela principal
janela = ctk.CTk()  # Criando a janela
janela.title('WorkFlowBot')  # Titulo da janela
janela.geometry('400x200')  # Definindo o tamanho da janela
janela.maxsize(width=400, height=200)  # Definindo o tamanho maximo da tela
janela.minsize(width=400, height=200)  # Definindo o tamanho minimo da tela
janela.resizable(width=False, height=False)  # Bloqueando ajuste de tela

# Rótulo para indicar ao usuário o que deve ser inserido
rotulo = ctk.CTkLabel(janela,
                      text="Insira o número da empresa:",
                      font=('Arial', 12))
rotulo.grid(row=0, column=0, columnspan=2, pady=10)

# # Campo de entrada personalizado
entry = ctk.CTkEntry(janela, font=('Arial', 12))
entry.grid(row=1, column=0, columnspan=2, padx=10)

# Botão para iniciar o bot
botao_extrair = ctk.CTkButton(master=janela,
                              corner_radius=10,
                              text='Extrair Dados',
                              command=extrair_dados,
                              fg_color='#109010',
                              hover_color='#176917')
botao_extrair.grid(row=2, column=0, padx=10, pady=10)

# Botão para fechar o bot
botao_fechar = ctk.CTkButton(master=janela,
                              corner_radius=10,
                              text='Fechar Bot',
                              command=janela.destroy,
                              fg_color='#fc4b08',
                              hover_color='#b94317')
botao_fechar.grid(row=2, column=1, padx=10, pady=10)

# Configurando o redimensionamento das colunas
janela.columnconfigure(0, weight=1)  # Expandir a primeira coluna
janela.columnconfigure(1, weight=1)  # Expandir a segunda coluna

# Iniciar o loop principal da interface gráfica
janela.mainloop()
