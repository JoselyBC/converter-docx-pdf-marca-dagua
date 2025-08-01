import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
import fitz  # PyMuPDF
import os
import shutil
import tempfile
# from docx2pdf import convert  <-- REMOVIDO
import comtypes.client       # <-- NOVA IMPORTAÇÃO
import traceback
import gc
import time
import uuid

# --- NOVA FUNÇÃO DE CONVERSÃO ---
def converter_docx_para_pdf(caminho_docx, caminho_pdf):
    """
    Converte um único arquivo DOCX para PDF usando a biblioteca comtypes.
    Esta função é mais robusta que a docx2pdf em alguns sistemas.
    """
    # Constante para o formato PDF no Word
    wdFormatPDF = 17
    word = None
    doc = None
    
    # Paths precisam ser absolutos para o COM
    caminho_docx_abs = os.path.abspath(caminho_docx)
    caminho_pdf_abs = os.path.abspath(caminho_pdf)

    try:
        # Inicia o Word em segundo plano
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False # Não mostrar a janela do Word
        
        # Abre o documento
        print(f"Abrindo documento: {caminho_docx_abs}")
        doc = word.Documents.Open(caminho_docx_abs)
        
        # Salva como PDF
        print(f"Salvando como PDF em: {caminho_pdf_abs}")
        doc.SaveAs(caminho_pdf_abs, FileFormat=wdFormatPDF)
        
    except Exception as e:
        # Levanta a exceção para ser tratada no loop principal
        raise e
    finally:
        # GARANTE que o Word seja fechado, mesmo se ocorrer um erro
        if doc:
            doc.Close()
            print("Documento do Word fechado.")
        if word:
            word.Quit()
            print("Aplicação Word encerrada.")
        # Libera os objetos COM
        doc = None
        word = None
        gc.collect()

# --- Funções de seleção (sem alterações) ---
def selecionar_logo():
    global g_caminho_logo
    caminho = filedialog.askopenfilename(
        title="Selecione a Imagem da Marca d'Água",
        filetypes=[("Imagem PNG", "*.png"), ("Imagem JPEG", "*.jpg;*.jpeg"), ("Todos os arquivos", "*.*")]
    )
    if caminho:
        g_caminho_logo = caminho
        entrada_logo.delete(0, tk.END)
        entrada_logo.insert(0, g_caminho_logo)

def selecionar_docx():
    global g_caminhos_docx_str
    caminhos_tupla = filedialog.askopenfilenames(
        title="Selecione os Arquivos DOCX",
        filetypes=[("Documento Word", "*.docx"), ("Todos os arquivos", "*.*")]
    )
    if caminhos_tupla:
        g_caminhos_docx_str = ";".join(caminhos_tupla)
        entrada_docx.delete(0, tk.END)
        entrada_docx.insert(0, g_caminhos_docx_str)

def validar_transparencia(valor_str):
    try:
        valor_int = int(valor_str)
        if 0 <= valor_int <= 255: return True
        else: messagebox.showerror("Erro de Validação", "A transparência deve ser um número inteiro entre 0 e 255."); return False
    except ValueError:
        messagebox.showerror("Erro de Validação", "A transparência deve ser um número inteiro válido."); return False

def tentar_remover_arquivo(caminho_arquivo, tentativas=3, delay=0.5):
    for i in range(tentativas):
        try:
            if os.path.exists(caminho_arquivo): os.remove(caminho_arquivo); return True
        except Exception: time.sleep(delay)
    return False

# --- Função principal modificada ---
def converter_arquivos():
    caminho_logo_atual = entrada_logo.get()
    caminhos_docx_str_atual = entrada_docx.get()
    transparencia_str = entrada_transparencia.get()

    if not caminho_logo_atual or not caminhos_docx_str_atual:
        messagebox.showerror("Erro", "Por favor, selecione a imagem da marca d'água e pelo menos um arquivo DOCX."); return

    if not validar_transparencia(transparencia_str): return
    
    transparencia_int = int(transparencia_str)
    lista_arquivos_docx = [p.strip() for p in caminhos_docx_str_atual.split(";") if p.strip()]

    if not lista_arquivos_docx:
        messagebox.showerror("Erro", "Nenhum arquivo DOCX válido selecionado."); return

    total_arquivos = len(lista_arquivos_docx)
    progresso_bar['maximum'] = total_arquivos
    progresso_bar['value'] = 0
    janela.update_idletasks()

    arquivos_convertidos_sucesso = 0
    erros_ocorridos = []

    pasta_temp_docx2pdf = os.path.join(tempfile.gettempdir(), f"conversor_pdf_temp_{uuid.uuid4().hex[:8]}")
    os.makedirs(pasta_temp_docx2pdf, exist_ok=True)

    for i, caminho_docx_original in enumerate(lista_arquivos_docx):
        doc_pdf_pymupdf = None
        caminho_pdf_intermediario = ""

        try:
            nome_base_docx = os.path.basename(caminho_docx_original)
            nome_pdf_interm = f"{os.path.splitext(nome_base_docx)[0]}_{uuid.uuid4().hex[:6]}.pdf"
            caminho_pdf_intermediario = os.path.join(pasta_temp_docx2pdf, nome_pdf_interm)
            
            # <<< MUDANÇA PRINCIPAL AQUI >>>
            converter_docx_para_pdf(caminho_docx_original, caminho_pdf_intermediario)
            
            if not os.path.exists(caminho_pdf_intermediario):
                raise FileNotFoundError(f"PDF intermediário não foi criado. Verifique se o MS Word está instalado.")
            
            gc.collect()
            time.sleep(0.5) 

            doc_pdf_pymupdf = fitz.open(caminho_pdf_intermediario)

            if doc_pdf_pymupdf.page_count == 0:
                raise ValueError("O documento está vazio ou não pôde ser lido corretamente.")

            with tempfile.TemporaryDirectory(prefix="logo_pagina_") as temp_dir_logos_pagina:
                for page_num, page_obj in enumerate(doc_pdf_pymupdf):
                    largura_pagina = int(page_obj.rect.width)
                    altura_pagina = int(page_obj.rect.height)
                    img_logo_pil = Image.open(caminho_logo_atual).convert("RGBA")
                    r, g, b, a = img_logo_pil.split()
                    a = a.point(lambda p: int(p * (int(transparencia_int) / 255.0)))
                    img_logo_pil_transp = Image.merge('RGBA', (r, g, b, a))
                    img_logo_redim = img_logo_pil_transp.resize((largura_pagina, altura_pagina), Image.Resampling.LANCZOS)
                    caminho_logo_pagina_temp = os.path.join(temp_dir_logos_pagina, f"logo_temp_p{page_num}.png")
                    img_logo_redim.save(caminho_logo_pagina_temp)
                    rect_logo = fitz.Rect(0, 0, largura_pagina, altura_pagina)
                    page_obj.insert_image(rect_logo, filename=caminho_logo_pagina_temp, overlay=True)
                    texto_rodape = "Escola da Nuvem — Todos os direitos reservados."
                    ponto_rodape = fitz.Point(50, altura_pagina - 30) 
                    page_obj.insert_text(ponto_rodape, texto_rodape, fontsize=8, fontname="helv", color=(0.33, 0.33, 0.33))

            pasta_saida_pdf = os.path.dirname(caminho_docx_original)
            nome_pdf_final_base = os.path.splitext(nome_base_docx)[0] + '.pdf'
            caminho_pdf_final_completo = os.path.join(pasta_saida_pdf, nome_pdf_final_base)
            
            doc_pdf_pymupdf.save(caminho_pdf_final_completo)
            arquivos_convertidos_sucesso +=1

        except Exception as e:
            error_message = str(e)
            print(f"ERRO no processamento de {caminho_docx_original}: {error_message}")
            print(traceback.format_exc())
            erros_ocorridos.append(f"{os.path.basename(caminho_docx_original)}: {error_message}")
        
        finally:
            if doc_pdf_pymupdf: doc_pdf_pymupdf.close(); doc_pdf_pymupdf = None
            gc.collect()
            time.sleep(0.5) 
            if caminho_pdf_intermediario and os.path.exists(caminho_pdf_intermediario):
                tentar_remover_arquivo(caminho_pdf_intermediario, tentativas=5, delay=0.7)
            progresso_bar['value'] = i + 1
            janela.update_idletasks()

    try:
        if os.path.exists(pasta_temp_docx2pdf): shutil.rmtree(pasta_temp_docx2pdf)
    except Exception as e:
        messagebox.showwarning("Aviso de Limpeza", f"Não foi possível remover todos os arquivos temporários em:\n{pasta_temp_docx2pdf}\n\nPode ser necessário removê-los manualmente.")

    if arquivos_convertidos_sucesso == total_arquivos:
        messagebox.showinfo("Sucesso", f"Todos os {total_arquivos} arquivos foram convertidos com sucesso!")
    else:
        msg_erro_final = f"{arquivos_convertidos_sucesso} de {total_arquivos} arquivos convertidos.\n"
        if erros_ocorridos: msg_erro_final += "Erros:\n" + "\n".join(erros_ocorridos)
        messagebox.showwarning("Concluído com Erros", msg_erro_final)

# --- Interface Gráfica (Tkinter) - Sem alterações ---
janela = tk.Tk()
janela.title("Conversor DOCX para PDF com Marca d'Água Personalizada")
janela.geometry("600x450")
frame_logo = tk.Frame(janela)
frame_logo.pack(pady=5, fill=tk.X, padx=10)
tk.Label(frame_logo, text="Imagem da Marca d'Água:").pack(side=tk.LEFT)
entrada_logo = tk.Entry(frame_logo, width=50)
entrada_logo.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5,0))
tk.Button(frame_logo, text="Selecionar...", command=selecionar_logo).pack(side=tk.LEFT, padx=(5,0))
frame_transp = tk.Frame(janela)
frame_transp.pack(pady=5, fill=tk.X, padx=10)
tk.Label(frame_transp, text="Transparência (0-255, 0=totalmente transparente):").pack(side=tk.LEFT)
entrada_transparencia = tk.Entry(frame_transp, width=10)
entrada_transparencia.insert(0, "15")
entrada_transparencia.pack(side=tk.LEFT, padx=(5,0))
frame_docx = tk.Frame(janela)
frame_docx.pack(pady=5, fill=tk.X, padx=10)
tk.Label(frame_docx, text="Arquivos DOCX (selecione um ou mais):").pack(side=tk.LEFT)
entrada_docx = tk.Entry(frame_docx, width=50)
entrada_docx.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5,0))
tk.Button(frame_docx, text="Selecionar...", command=selecionar_docx).pack(side=tk.LEFT, padx=(5,0))
btn_converter = tk.Button(janela,text="Converter para PDF com Marca d'Água",command=converter_arquivos,bg="green",fg="white",font=("Arial", 10, "bold"))
btn_converter.pack(pady=20, ipady=5)
progresso_label = tk.Label(janela, text="Progresso:")
progresso_label.pack()
progresso_bar = ttk.Progressbar(janela, orient='horizontal', length=400, mode='determinate')
progresso_bar.pack(pady=10, padx=10, fill=tk.X)
janela.mainloop()