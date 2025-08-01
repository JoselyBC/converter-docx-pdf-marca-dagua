# Conversor de PDF com Marca d'Água

Este é um aplicativo de desktop desenvolvido para facilitar a conversão de documentos do Word (`.docx`) para o formato PDF, adicionando automaticamente uma marca d'água personalizada e um rodapé em todas as páginas.

## 🚀 Tecnologias Usadas

Este projeto foi construído utilizando as seguintes tecnologias e bibliotecas:

*   **Python:** Linguagem principal do projeto (recomendado versão 3.11 para máxima compatibilidade).
*   **Tkinter:** Para a construção da interface gráfica do usuário (GUI).
*   **Pillow (PIL):** Para o processamento e ajuste de transparência da imagem da marca d'água.
*   **PyMuPDF (fitz):** Para a manipulação do arquivo PDF, inserindo a marca d'água e o texto de rodapé.
*   **comtypes:** Para a comunicação direta e robusta com o Microsoft Word, realizando a conversão do DOCX para PDF.
*   **PyInstaller:** Para empacotar a aplicação em um único arquivo executável (`.exe`) para fácil distribuição no Windows.

## 📋 O que o projeto faz

A aplicação permite ao usuário:

*   **Converter Arquivos em Lote:** Selecionar um ou múltiplos documentos `.docx` para serem convertidos de uma só vez.
*   **Aplicar Marca d'Água:** Escolher uma imagem (como um logo) para ser aplicada como marca d'água em todas as páginas dos documentos gerados.
*   **Ajustar Transparência:** Controlar o nível de transparência da marca d'água para um resultado mais sutil ou mais forte.
*   **Personalizar o Rodapé:** Definir um texto personalizado para ser inserido no rodapé de cada página, ou deixá-lo em branco para não incluir rodapé.
*   **Interface Simples:** Oferece uma janela intuitiva para que usuários sem conhecimento técnico possam realizar as conversões facilmente.

## 💻 Como Usar a Aplicação (Para Usuários Finais)

Esta seção é para quem vai apenas utilizar o programa pronto (`.exe`).

#### Pré-requisito

*   É essencial ter o **Microsoft Word** instalado no seu computador, pois ele é utilizado no processo de conversão.

#### Instruções

1.  Você receberá ou fará o download de um único arquivo chamado **`conversor_pdf.exe`**.
2.  Crie uma nova pasta em seu computador para organizar seus arquivos e mova o `.exe` para dentro dela.
3.  Dê um duplo-clique no arquivo `conversor_pdf.exe` para iniciar o programa.
4.  Na janela da aplicação:
    *   Clique em "Selecionar..." para escolher a imagem da marca d'água.
    *   **(Opcional)** Edite o texto no campo "Texto do Rodapé" para personalizar o rodapé.
    *   Clique em "Selecionar..." para escolher os arquivos `.docx` que deseja converter.
    *   Clique no botão verde **"Converter para PDF com Marca d'Água"**.
5.  Aguarde o processo terminar. Os arquivos PDF convertidos serão salvos na mesma pasta onde os arquivos `.docx` originais estão localizados.

## 🔧 Para Desenvolvedores (Rodando e Compilando o Projeto)

Esta seção é para quem deseja rodar o código-fonte e gerar o executável por conta própria.

#### 1. Preparação do Ambiente

*   Garanta que você tenha o **Python 3.11** instalado (para máxima compatibilidade com as bibliotecas de automação do Windows).
*   Tenha o **Microsoft Word** instalado.
*   Clone ou baixe o repositório para o seu computador.
*   Abra um terminal na pasta do projeto e crie um ambiente virtual:
    ```bash
    # Use o comando específico de versão para garantir que está usando a versão correta
    py -3.11 -m venv venv
    ```
*   Ative o ambiente virtual:
    ```bash
    # No Windows
    .\venv\Scripts\Activate.ps1
    ```

#### 2. Instalação das Dependências

Com o ambiente virtual ativo, instale todas as bibliotecas necessárias:
```bash
pip install Pillow PyMuPDF comtypes pyinstaller
```

#### 3. Rodando o Script

Para testar a aplicação antes de compilá-la, execute:
```bash
python converte-docx-pdf.py
```

#### 4. Gerando o Executável

Para compilar o projeto em um único arquivo `.exe`:
```bash
pyinstaller --onefile --windowed --clean converte-docx-pdf.py
```
*   Após o processo terminar, o arquivo executável **`conversor_pdf.exe`** estará localizado dentro da pasta **`dist`** que foi criada.

---

### Projeto desenvolvido por **Josely Castro**.

[<img src="https://img.shields.io/badge/linkedin-%230077B5.svg?&style=for-the-badge&logo=linkedin&logoColor=white" />](https://www.linkedin.com/in/joselybcastro/) [<img src="https://img.shields.io/badge/github-%23121011.svg?&style=for-the-badge&logo=github&logoColor=white" />](https://github.com/joselyBC)