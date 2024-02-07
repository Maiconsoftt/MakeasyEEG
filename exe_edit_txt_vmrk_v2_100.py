import os
import time
import pandas as pd
import tkinter as tk
from tkinter import filedialog, OptionMenu, messagebox, ttk
from PIL import Image, ImageTk

def loading_window(files):
    loading_window = tk.Toplevel()
    loading_window.title("Processando arquivos...")

    # Barra de progresso
    progress_label = tk.Label(loading_window, text="Processando arquivo:")
    progress_label.pack()

    progress_bar = ttk.Progressbar(loading_window, orient="horizontal", length=300, mode="determinate")
    progress_bar.pack(pady=10)

    # Exibir os nomes dos arquivos rapidamente
    for file_name in files:
        progress_bar.step(100 / len(files))  # Atualiza a barra de progresso
        progress_label.config(text=file_name)  # Atualiza o nome do arquivo
        loading_window.update()
        time.sleep(0.2)  # Pausa a execução por 0.2 segundos

    # Fecha a janela de carregamento
    loading_window.destroy()

def obter_caminho_pasta():
    pasta = filedialog.askdirectory(title="Selecione a pasta contendo os arquivos markers")
    entry_pasta.delete(0, tk.END)
    entry_pasta.insert(0, pasta)

def obter_caminho_excel():
    arquivo_excel = filedialog.askopenfilename(title="Selecione o arquivo Excel")
    entry_excel.delete(0, tk.END)
    entry_excel.insert(0, arquivo_excel)

def processar_dados(dir_path, excel_path, extensao_arquivo):
    # Verifica se a pasta contém arquivos com a extensão selecionada
    files = [file for file in os.listdir(dir_path) if file.endswith(extensao_arquivo)]

    if not files:
        messagebox.showerror("Erro!", "A pasta selecionada não contém arquivos com a extensão especificada.")
        return
    else:
        message = f"Foram encontrados {len(files)} arquivos:\n" + "\n".join(files)
        messagebox.showinfo("Arquivos Encontrados", message)

    # Carrega os novos valores do arquivo Excel para um DataFrame
    df_excel = pd.read_excel(excel_path)

    # Mapeamento de arquivos para suas respectivas colunas no Excel
    file_column_mapping = {file.replace(extensao_arquivo, ''): coluna for file, coluna in zip(files, df_excel.columns)}

    message = "Os arquivos foram associados às seguintes colunas do excel:\n"
    for file_name, coluna_excel in file_column_mapping.items():
        message += f"{file_name}{extensao_arquivo}: {coluna_excel}\n"
    messagebox.showinfo("Mapeamento de arquivos e colunas", message)

    # Defina as linhas a serem editadas
    try:
        linha_inicial = int(entry_linha_inicial.get())
        linha_final = int(entry_linha_final.get())
    except ValueError:
        messagebox.showerror("Erro!", "Por favor, insira valores numéricos para as linhas inicial e final.")
        return

    valor_procurado = entry_valor_procurado.get()
    if not valor_procurado.strip():  # Verifica se o valor procurado está em branco
        messagebox.showerror("Erro!", "Por favor, insira um valor para procurar nos arquivos.")
        return

    # Verifica se a extensão selecionada é compatível com a extensão dos arquivos na pasta
    extensoes_arquivos = {os.path.splitext(file)[1] for file in files}
    if extensao_arquivo not in extensoes_arquivos:
        messagebox.showerror("Erro!", "A extensão de arquivo selecionada não corresponde à extensão dos arquivos na pasta.")
        messagebox.showerror("Verifique a pasta", f"Selecione uma pasta que contenha arquivos com extensões compatíveis: {', '.join(extensoes_arquivos)}")
        return
    
    loading_window (files)

    # Itera sobre o mapeamento de arquivo e coluna
    for file_name, coluna_excel in file_column_mapping.items():
        # Constrói o caminho completo para o arquivo
        file_path = os.path.join(dir_path, f"{file_name}{extensao_arquivo}")

        # Lista para armazenar as linhas modificadas    
        updated_lines = []
        
        if extensao_arquivo == '.txt':
            novos_valores_formatados = ['     ' + str(valor) for valor in df_excel[coluna_excel].tolist()]
        else:
            # Carrega a coluna específica do Excel para esse arquivo
            novos_valores_formatados = ['' + str(valor) for valor in df_excel[coluna_excel].tolist()]

        # Abre o arquivo em modo de leitura ('r')
        with open(file_path, 'r') as file:
            # Leia cada linha do arquivo
            for linha_numero, line in enumerate(file, start=1):
                # Verifique se a linha está dentro do intervalo desejado
                if linha_inicial <= linha_numero <= linha_final:
                    # Faça a edição apenas nas linhas desejadas
                    # Procura e substitui o valor especificado em cada linha
                    updated_line = line.replace(valor_procurado, novos_valores_formatados[linha_numero - linha_inicial])
                else:
                    updated_line = line

                updated_lines.append(updated_line)

        # Abre o arquivo em modo de escrita ('w') para sobrescrever o conteúdo
        with open(file_path, 'w') as file:
            # Escreva as linhas modificadas de volta para o arquivo
            file.writelines(updated_lines)

        #print(f"Alterações salvas com sucesso para o arquivo: {file_name}")

    messagebox.showinfo("Concluído", f"Todos os {len(files)} arquivos foram processados!\nFaça bom uso desses dados e que os deuses da pesquisa estejam convosco.")

# Criar a janela principal
root = tk.Tk()
root.title("Makeasy EGG - Editor de Markers (v1.75) - by onurB")

icone = Image.open("C:\\Users\\bruno_mwax2\\Documents\\0. Scripts\\Logo_gpmc.ico")
imagem = ImageTk.PhotoImage(icone)

root.tk.call('wm', 'iconphoto', root._w, imagem)

# Variável de controle para o menu suspenso
extensao_arquivo_var = tk.StringVar(root)
extensao_arquivo_var.set('.vmrk')  # Valor padrão

# Frame para agrupar o botão "Selecionar Pasta" e o menu suspenso
frame_menu = tk.Frame(root)
frame_menu.grid(row=3, column=0, padx=10, pady=5, sticky='n')

# Widgets e Layout
tk.Label(root, text="Selecione a pasta contendo os arquivos .txt ou .vmrk:", width=65).grid(column=0, row=1, pady=4)
entry_pasta = tk.Entry(root, width=64)
entry_pasta.grid(column=0, row=2)
tk.Button(frame_menu, text="Selecionar Pasta", command=obter_caminho_pasta).grid(row=0, column=0)  # Adicionando o botão ao frame

tk.Label(root, text="Selecione o arquivo Excel:").grid(column=0, row=7, pady=4)
entry_excel = tk.Entry(root, width=64)
entry_excel.grid(column=0, row=8)
tk.Button(root, text="Selecionar Excel", command=obter_caminho_excel).grid(column=0, row=9, pady=3)

tk.Label(root, text='Linha inicial dos arquivos com os marcadores:').grid(column=0, row=20, pady=4)
entry_linha_inicial = tk.Entry(root, width=10)
entry_linha_inicial.grid(column=0, row=21, pady=3)

tk.Label(root, text='Linha final dos arquivos com os marcadores:').grid(column=0, row=25, pady=4)
entry_linha_final = tk.Entry(root, width=10)
entry_linha_final.grid(column=0, row=26, pady=3)

tk.Label(root, text='Insira o texto procurado:').grid(column=0, row=30, pady=4)
entry_valor_procurado = tk.Entry(root, width=10)
entry_valor_procurado.grid(column=0, row=31, pady=3)

# Menu suspenso para selecionar a extensão do arquivo
extensao_menu = OptionMenu(frame_menu, extensao_arquivo_var, '.txt', '.vmrk')
extensao_menu.grid(row=0, column=1, padx=(5, 0))  # Adicionando o menu suspenso ao frame

tk.Button(root, text="Processar", command=lambda: processar_dados(entry_pasta.get(), entry_excel.get(), extensao_arquivo_var.get())).grid(column=0, row=45, pady=25, ipady=5, ipadx=10)

# Iniciar o loop principal
root.mainloop()