import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert

# Variáveis para armazenar os caminhos 
docx_path = ""
pdf_path = ""

def selecionar_arquivo():
    global docx_path
    docx_path = filedialog.askopenfilename(filetypes=[("Documentos Word", "*.docx")])

def selecionar_pasta():
    global pdf_path
    pdf_path = filedialog.askdirectory()

def converter():
    global docx_path, pdf_path

    # Verifica se os caminhos estão definidos
    if docx_path and pdf_path:
        # Construir o caminho do arquivo PDF
        pdf_file = pdf_path + "/" + docx_path.split("/")[-1].split(".")[0] + ".pdf"

        try:
            # Converter o arquivo
            convert(docx_path, pdf_file)

            # Mostrar mensagem de confirmação na interface
            mensagem_confirmacao = tk.Label(janela, text="Conversão realizada com sucesso!")
            mensagem_confirmacao.pack()
        except Exception as e:
            # Mostrar mensagem de erro na interface
            mensagem_erro = tk.Label(janela, text=f"Ocorreu um erro durante a conversão: {e}")
            mensagem_erro.pack()
    else:
        # Exibir mensagem de erro se algum caminho não estiver definido
        mensagem_erro = tk.Label(janela, text="Por favor, selecione um arquivo DOCX e uma pasta de destino.")
        mensagem_erro.pack()

# Criar a janela principal
janela = tk.Tk()
janela.title("Conversor DOCX para PDF")

# Rótulo para o campo de seleção do arquivo
label_arquivo = tk.Label(janela, text="Selecione o arquivo DOCX:")
label_arquivo.pack()

# Botão para selecionar o arquivo
botao_selecionar_arquivo = tk.Button(janela, text="Procurar", command=selecionar_arquivo)
botao_selecionar_arquivo.pack()

# Rótulo para o campo de seleção da pasta
label_pasta = tk.Label(janela, text="Selecione a pasta de destino:")
label_pasta.pack()

# Botão para selecionar a pasta
botao_selecionar_pasta = tk.Button(janela, text="Procurar", command=selecionar_pasta)
botao_selecionar_pasta.pack()

# Botão para iniciar a conversão
botao_converter = tk.Button(janela, text="Converter", command=converter)
botao_converter.pack()

janela.mainloop()