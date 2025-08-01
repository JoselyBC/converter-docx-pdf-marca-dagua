import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
import fitz  # PyMuPDF
import os
import shutil
import tempfile # Ainda usado para os logos temporários das páginas
from docx2pdf import convert
import traceback
import gc # Garbage Collector
import time # Para pausas
import uuid # Para nomes de arquivo únicos

# --- Variáveis globais e funções de seleção (como na versão anterior) ---
g_caminho_logo = ""
g_caminhos_docx_str = ""

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
        g_caminhos_docx_str = ";".join(caminhos_tupla) # Junta com ';' para o Entry
        entrada_docx.delete(0, tk.END)
        entrada_docx.insert(0, g_caminhos_docx_str)

def validar_transparencia(valor_str):
    try:
        valor_int = int(valor_str)
        if 0 <= valor_int <= 255:
            return True
        else:
            messagebox.showerror("Erro de Validação", "A transparência deve ser um número inteiro entre 0 e 255.")
            return False
    except ValueError:
        messagebox.showerror("Erro de Validação", "A transparência deve ser um número inteiro válido.")
        return False
# --- Fim das funções de seleção ---


def tentar_remover_arquivo(caminho_arquivo, tentativas=3, delay=0.5):
    """Tenta remover um arquivo com múltiplas tentativas e delays."""
    for i in range(tentativas):
        try:
            if os.path.exists(caminho_arquivo):
                os.remove(caminho_arquivo)
                print(f"Arquivo {caminho_arquivo} removido com sucesso na tentativa {i+1}.")
                return True
        except PermissionError:
            print(f"PermissionError ao tentar remover {caminho_arquivo} (tentativa {i+1}/{tentativas}). Aguardando {delay}s...")
            time.sleep(delay)
        except Exception as e:
            print(f"Erro inesperado ao tentar remover {caminho_arquivo} (tentativa {i+1}): {e}")
            time.sleep(delay) # Também esperar em outros erros
    print(f"Falha ao remover {caminho_arquivo} após {tentativas} tentativas.")
    return False


def converter_arquivos():
    caminho_logo_atual = entrada_logo.get()
    caminhos_docx_str_atual = entrada_docx.get()
    transparencia_str = entrada_transparencia.get()

    if not caminho_logo_atual or not caminhos_docx_str_atual:
        messagebox.showerror("Erro", "Por favor, selecione a imagem da marca d'água e pelo menos um arquivo DOCX.")
        return

    if not validar_transparencia(transparencia_str):
        return
    
    transparencia_int = int(transparencia_str)
    lista_arquivos_docx = [p.strip() for p in caminhos_docx_str_atual.split(";") if p.strip()]

    if not lista_arquivos_docx:
        messagebox.showerror("Erro", "Nenhum arquivo DOCX válido selecionado.")
        return

    total_arquivos = len(lista_arquivos_docx)
    progresso_bar['maximum'] = total_arquivos
    progresso_bar['value'] = 0
    janela.update_idletasks()

    arquivos_convertidos_sucesso = 0
    erros_ocorridos = []

    # Criar uma pasta temporária de nível superior para todos os PDFs intermediários
    # Isso nos dá mais controle sobre quando ela é deletada.
    pasta_temp_docx2pdf = os.path.join(tempfile.gettempdir(), f"conversor_pdf_temp_{uuid.uuid4().hex[:8]}")
    os.makedirs(pasta_temp_docx2pdf, exist_ok=True)
    print(f"Pasta temporária principal para PDFs intermediários: {pasta_temp_docx2pdf}")

    for i, caminho_docx_original in enumerate(lista_arquivos_docx):
        doc_pdf_pymupdf = None # Para PyMuPDF
        caminho_pdf_intermediario = "" # Caminho do PDF criado por docx2pdf

        try:
            nome_base_docx = os.path.basename(caminho_docx_original)
            
            # Caminho do PDF intermediário DENTRO da nossa pasta_temp_docx2pdf
            # Usar um nome único para evitar conflitos se o mesmo nome de arquivo for processado em rápida sucessão (improvável com GUI)
            nome_pdf_interm = f"{os.path.splitext(nome_base_docx)[0]}_{uuid.uuid4().hex[:6]}.pdf"
            caminho_pdf_intermediario = os.path.join(pasta_temp_docx2pdf, nome_pdf_interm)
            
            print(f"Convertendo {caminho_docx_original} para PDF intermediário em: {caminho_pdf_intermediario}")
            convert(caminho_docx_original, caminho_pdf_intermediario) # docx2pdf cria e (deveria) fechar o PDF
            
            if not os.path.exists(caminho_pdf_intermediario):
                raise FileNotFoundError(f"PDF intermediário {caminho_pdf_intermediario} não foi criado por docx2pdf.")
            print(f"PDF intermediário criado: {caminho_pdf_intermediario}")

            # Liberar memória e dar um tempo para o Word (se ele foi usado)
            gc.collect()
            time.sleep(0.5) 

            print(f"Abrindo {caminho_pdf_intermediario} com PyMuPDF")
            doc_pdf_pymupdf = fitz.open(caminho_pdf_intermediario)
            print(f"PyMuPDF abriu {caminho_pdf_intermediario}")

            # Usar um diretório temporário SEPARADO para os logos de página da PIL
            # Este TemporaryDirectory é para arquivos que SÃO controlados pelo nosso script.
            with tempfile.TemporaryDirectory(prefix="logo_pagina_") as temp_dir_logos_pagina:
                for page_num, page_obj in enumerate(doc_pdf_pymupdf):
                    largura_pagina = int(page_obj.rect.width)
                    altura_pagina = int(page_obj.rect.height)

                    img_logo_pil = Image.open(caminho_logo_atual).convert("RGBA")
                    r, g, b, a = img_logo_pil.split()
                    alpha_pil = int(transparencia_int) # 0 (transparente) a 255 (opaco) para PIL
                    a = a.point(lambda p: int(p * (alpha_pil / 255.0)))
                    img_logo_pil_transp = Image.merge('RGBA', (r, g, b, a))

                    img_logo_redim = img_logo_pil_transp.resize((largura_pagina, altura_pagina), Image.Resampling.LANCZOS)
                    
                    caminho_logo_pagina_temp = os.path.join(temp_dir_logos_pagina, f"logo_temp_p{page_num}.png")
                    img_logo_redim.save(caminho_logo_pagina_temp)

                    rect_logo = fitz.Rect(0, 0, largura_pagina, altura_pagina)
                    page_obj.insert_image(rect_logo, filename=caminho_logo_pagina_temp, overlay=True)
                    
                    texto_rodape = "Escola da Nuvem — Todos os direitos reservados."
                    ponto_rodape = fitz.Point(50, altura_pagina - 30) 
                    page_obj.insert_text(ponto_rodape, texto_rodape, fontsize=8, fontname="helv", color=(0.33, 0.33, 0.33))
                    # Não precisamos mais remover o caminho_logo_pagina_temp aqui, o TemporaryDirectory fará isso.
            # Fim do with temp_dir_logos_pagina

            pasta_saida_pdf = os.path.dirname(caminho_docx_original)
            nome_pdf_final_base = os.path.splitext(nome_base_docx)[0] + '.pdf'
            caminho_pdf_final_completo = os.path.join(pasta_saida_pdf, nome_pdf_final_base)
            
            print(f"Salvando PDF final em {caminho_pdf_final_completo}")
            doc_pdf_pymupdf.save(caminho_pdf_final_completo)
            print("PDF final salvo.")

            arquivos_convertidos_sucesso +=1

        except Exception as e:
            print(f"ERRO no processamento de {caminho_docx_original}: {e}")
            print(traceback.format_exc())
            erros_ocorridos.append(f"{os.path.basename(caminho_docx_original)}: {e}")
        
        finally:
            # GARANTIR que o PyMuPDF feche o arquivo intermediário
            if doc_pdf_pymupdf:
                print(f"Fechando PDF intermediário {caminho_pdf_intermediario} com PyMuPDF.")
                doc_pdf_pymupdf.close()
                doc_pdf_pymupdf = None
                print("PyMuPDF fechou o arquivo.")
            
            gc.collect() # Coleta de lixo
            time.sleep(0.5) # Dar um tempo extra

            # Tentar remover o PDF INTERMEDIÁRIO explicitamente
            if caminho_pdf_intermediario and os.path.exists(caminho_pdf_intermediario):
                print(f"Tentando remover o PDF intermediário: {caminho_pdf_intermediario}")
                tentar_remover_arquivo(caminho_pdf_intermediario, tentativas=5, delay=0.7)
            
            progresso_bar['value'] = i + 1
            janela.update_idletasks()
            print("-" * 30)

    # Após o loop, tentar remover a pasta temporária principal
    print(f"Tentando remover a pasta temporária principal: {pasta_temp_docx2pdf}")
    try:
        if os.path.exists(pasta_temp_docx2pdf):
            shutil.rmtree(pasta_temp_docx2pdf)
            print(f"Pasta temporária principal {pasta_temp_docx2pdf} removida.")
    except Exception as e:
        print(f"ERRO ao remover pasta temporária principal {pasta_temp_docx2pdf}: {e}")
        messagebox.showwarning("Aviso de Limpeza", f"Não foi possível remover automaticamente todos os arquivos temporários em:\n{pasta_temp_docx2pdf}\n\nPode ser necessário removê-los manualmente.")


    if arquivos_convertidos_sucesso == total_arquivos:
        messagebox.showinfo("Sucesso", f"Todos os {total_arquivos} arquivos foram convertidos com sucesso!")
    else:
        msg_erro_final = f"{arquivos_convertidos_sucesso} de {total_arquivos} arquivos convertidos.\n"
        if erros_ocorridos:
            msg_erro_final += "Erros:\n" + "\n".join(erros_ocorridos)
        messagebox.showwarning("Concluído com Erros", msg_erro_final)

    print("Função converter_arquivos finalizada.")


# --- Interface Gráfica (Tkinter) ---
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

btn_converter = tk.Button(
    janela,
    text="Converter para PDF com Marca d'Água",
    command=converter_arquivos,
    bg="green",
    fg="white",
    font=("Arial", 10, "bold")
)
btn_converter.pack(pady=20, ipady=5)

progresso_label = tk.Label(janela, text="Progresso:")
progresso_label.pack()
progresso_bar = ttk.Progressbar(janela, orient='horizontal', length=400, mode='determinate')
progresso_bar.pack(pady=10, padx=10, fill=tk.X)

janela.mainloop()