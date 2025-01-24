import json
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import pandas as pd
from datetime import datetime

"""
Automação para filtrar linhas de um arquivo Excel e gerar um arquivo JSON com os eventos futuros.

"""

# Variável global para controlar o cancelamento do processamento
cancelar = False


# def filtrar_linhas_por_data_futura(caminho_excel, termos_busca):
#     """Filtra linhas do Excel com base nos termos de busca na segunda coluna
#     e identifica datas futuras em 2025 na quarta coluna.

#     Args:
#         caminho_excel (str): Caminho para o arquivo Excel.
#         termos_busca (list): Lista de termos para buscar na segunda coluna.

#     Returns:
#         pd.DataFrame: DataFrame contendo as linhas que atendem aos critérios.
#     """
#     # Leitura do arquivo Excel
#     df = pd.read_excel(caminho_excel)

#     # Verificar se há ao menos quatro colunas
#     if df.shape[1] < 4:
#         raise ValueError("O arquivo Excel precisa ter ao menos quatro colunas.")

#     # Filtrar linhas com os termos de busca na segunda coluna
#     termos_regex = '|'.join(map(str.lower, termos_busca))
#     linhas_filtradas = df[df.iloc[:, 1].astype(str).str.lower().str.contains(termos_regex, na=False)]

#     # Verificar se as datas na quarta coluna são futuras em 2025
#     datas_futuras_2025 = []
#     hoje = datetime.now()

#     for index, row in linhas_filtradas.iterrows():
#         try:
#             data = pd.to_datetime(row.iloc[3], errors='coerce')  # Converter valores da quarta coluna em datas
#             if data and data.year == 2025 and data > hoje:
#                 datas_futuras_2025.append(row)
#         except Exception:
#             continue  # Ignorar erros de conversão

#     # Criar DataFrame com as linhas que atendem ao critério
#     df_futuras_2025 = pd.DataFrame(datas_futuras_2025, columns=linhas_filtradas.columns)
    
#     return df_futuras_2025
        
def gerar_json_eventos(caminho_excel, termos_busca, caminho_saida):
    """Processa o Excel e gera um arquivo JSON com os eventos filtrados.

    Args:
        caminho_excel (str): Caminho para o arquivo Excel.
        termos_busca (list): Lista de termos para buscar na segunda coluna.
        caminho_saida (str): Caminho para salvar o arquivo JSON.
    """
    # Leitura do arquivo Excel
    df = pd.read_excel(caminho_excel)

    # Verificar se há ao menos quatro colunas
    if df.shape[1] < 4:
        raise ValueError("O arquivo Excel precisa ter ao menos quatro colunas.")

    # Filtrar linhas com os termos de busca na segunda coluna
    termos_regex = '|'.join(map(str.lower, termos_busca))
    linhas_filtradas = df[df.iloc[:, 1].astype(str).str.lower().str.contains(termos_regex, na=False)]

    # Verificar as datas na quarta coluna e coletar eventos futuros
    eventos = []
    hoje = datetime.now()
    for _, row in linhas_filtradas.iterrows():
        try:
            data = pd.to_datetime(row.iloc[3], errors='coerce')  # Converter para data
            if data and data.year == 2025 and data > hoje:
                evento = {
                    "nome": row.iloc[1],
                    "data": data.strftime("%Y-%m-%dT%H:%M:%S"),  # Formato ISO 8601
                }
                eventos.append(evento)
        except Exception as e:
            print(f"Erro ao processar linha: {e}")

    # Salvar eventos como JSON
    with open(caminho_saida, 'w', encoding='utf-8') as arquivo_json:
        json.dump(eventos, arquivo_json, ensure_ascii=False, indent=4)
    print(f"Arquivo JSON gerado com sucesso: {caminho_saida}")

def selecionar_destino_json():
    """Função para selecionar o caminho para salvar o arquivo JSON."""
    arquivo = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=(("Arquivos JSON", "*.json"), ("Todos os arquivos", "*.*")),
        title="Salvar arquivo JSON"
    )
    if arquivo:  # Verifica se o usuário selecionou um arquivo
        entrada_json.config(state='normal')  # Habilita a entrada para edição temporária
        entrada_json.delete(0, tk.END)
        entrada_json.insert(0, arquivo)
        entrada_json.config(state='readonly')  # Torna a entrada novamente somente leitura
        caminho_json.set(arquivo)  # Atualiza o valor do caminho_json


# Função para selecionar o arquivo Excel
def selecionar_excel():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*"))
    )
    if arquivo:
        caminho_excel.set(arquivo)
        atualizar_log(f"Arquivo Excel selecionado: {arquivo}")

# Função para iniciar o processamento dos dados
def iniciar_processamento():
    global cancelar
    cancelar = False
    
    excel = caminho_excel.get()
    json = caminho_json.get()
    
    if excel:
        atualizar_log("Iniciando processamento...")
        botao_iniciar.config(state=tk.DISABLED)  # Desabilitar o botão para evitar múltiplos cliques
        thread = threading.Thread(target=processar_dados, args=(excel, json))
        thread.start()
    else:
        messagebox.showwarning("Atenção", "Por favor, selecione o arquivo Excel.")

# Função do seu código que já está pronta para processar os dados
def processar_dados(excel, caminho_json):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return     
    # Arquivo Excel
    caminho_excel = excel
    caminho_saida_json = caminho_json
    # Search terms
    search_terms = [
        "prorrogações",
        "cobrança",
        "certificado digital"
    ]   
    resultado = gerar_json_eventos(caminho_excel, search_terms, caminho_saida_json)
    # resultado.to_excel("resultado_filtrado.xlsx", index=False)
    atualizar_log("Processamento concluído!", cor="verde")
    
# Função para cancelar o processamento
def cancelar_processamento():
    global cancelar
    cancelar = True
    atualizar_log("Cancelando processamento...")
    botao_fechar.config(state=tk.NORMAL)  # Habilitar o botão de fechar o programa

# Função para cancelar e fechar o programa
def fechar_programa():
    janela.quit()

# Função para finalizar o programa com uma mensagem
def finalizar_programa():
    messagebox.showinfo("Processo Finalizado", "O processamento foi concluído com sucesso!")
    botao_fechar.config(state=tk.NORMAL)  # Habilitar o botão de fechar o programa
    botao_iniciar.config(state=tk.NORMAL)  # Reabilitar o botão de iniciar

# Função para atualizar o log na área de texto
def atualizar_log(mensagem, cor=None):
    log_text.config(state=tk.NORMAL)  # Habilitar edição temporária
    if cor == "vermelho":
        log_text.insert(tk.END, mensagem + "\n", "vermelho")  # Inserir nova mensagem com tag 'vermelho'
    elif cor == "verde":
        log_text.insert(tk.END, mensagem + "\n", "verde")  # Inserir nova mensagem com tag 'verde'
    elif cor == "azul":
        log_text.insert(tk.END, mensagem + "\n", "azul")
    else:
        log_text.insert(tk.END, mensagem + "\n")  # Inserir nova mensagem sem tag
    log_text.config(state=tk.DISABLED)  # Desabilitar edição novamente
    log_text.see(tk.END)  # Scroll automático para a última linha

# Função para configurar a tag de cor no log
def configurar_tags_log():
    log_text.tag_config("vermelho", foreground="red")  # Configura a cor vermelha para a tag 'vermelho'
    log_text.tag_config("verde", foreground="green")  # Configura a cor vermelha para a tag 'verde'
    log_text.tag_config("azul", foreground="blue") # Configura a cor azul para a tag 'azul'
# Função main para encapsular a lógica do programa
def main():
    global janela, caminho_pasta, caminho_excel, caminho_json, botao_fechar, botao_iniciar, log_text, entrada_json

    

    # Criar a janela principal
    janela = tk.Tk()
    janela.title("Envio Mensagem Onvio")
    janela.geometry("600x400")
    janela.resizable(False, False)

    # Variáveis para armazenar os caminhos
    caminho_pasta = tk.StringVar()
    caminho_excel = tk.StringVar()
    caminho_json = tk.StringVar()
    
    # Frame para seleção de pasta e arquivo
    frame_selecao = tk.Frame(janela)
    frame_selecao.pack(pady=10)


    # Label e Botão para selecionar o arquivo Excel
    label_excel = tk.Label(frame_selecao, text="Arquivo Excel:")
    label_excel.grid(row=1, column=0, pady=5, padx=5)

    entrada_excel = tk.Entry(frame_selecao, textvariable=caminho_excel, width=50, state='readonly')
    entrada_excel.grid(row=1, column=1, padx=5)

    botao_excel = tk.Button(frame_selecao, text="Selecionar Excel", command=selecionar_excel)
    botao_excel.grid(row=1, column=2, padx=5)

     # Label e Botão para selecionar o destino do JSON
    label_json = tk.Label(frame_selecao, text="Arquivo JSON:")
    label_json.grid(row=2, column=0, pady=5, padx=5)

    entrada_json = tk.Entry(frame_selecao, textvariable=caminho_json, width=50, state='readonly')
    entrada_json.grid(row=2, column=1, padx=5)

    botao_json = tk.Button(frame_selecao, text="Salvar JSON", command=selecionar_destino_json)
    botao_json.grid(row=2, column=2, padx=5)
    
    
    # Botão para iniciar o processamento
    botao_iniciar = tk.Button(janela, text="Iniciar Processamento", command=iniciar_processamento)
    botao_iniciar.pack(pady=10)

    # Botão para cancelar e fechar o programa
    botao_cancelar = tk.Button(janela, text="Cancelar Processamento", command=cancelar_processamento)
    botao_cancelar.pack(pady=5)

    # Botão para fechar o programa (desabilitado até o processamento terminar)
    botao_fechar = tk.Button(janela, text="Fechar Programa", command=fechar_programa, state=tk.DISABLED)
    botao_fechar.pack(pady=5)

    # Frame para o log
    frame_log = tk.Frame(janela)
    frame_log.pack(pady=10, fill=tk.BOTH, expand=True)

    # Área de texto para o log com barra de rolagem
    log_text = scrolledtext.ScrolledText(frame_log, wrap=tk.WORD, height=10, state=tk.DISABLED)
    log_text.pack(fill=tk.BOTH, expand=True)

    # Configurar as tags de cor para o log
    configurar_tags_log()
    
    # Iniciar o loop da interface
    janela.mainloop()

# Garantir que o código só execute se este arquivo for o principal
if __name__ == '__main__':
    main()
