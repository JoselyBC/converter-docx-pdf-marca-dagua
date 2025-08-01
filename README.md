# Conversor de PDF com Marca d'Água

Este é um aplicativo de desktop desenvolvido para facilitar a conversão de documentos do Word (`.docx`) para o formato PDF, adicionando automaticamente uma marca d'água personalizada e um rodapé em todas as páginas.


*(Sugestão: Substitua o link acima pelo link de uma imagem do seu projeto no GitHub para que ela apareça aqui)*

## 🚀 Tecnologias Usadas

Este projeto foi construído utilizando as seguintes tecnologias e bibliotecas:

*   **Python:** Linguagem principal do projeto.
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
*   **Adicionar Rodapé Padrão:** Insere automaticamente o texto "Escola da Nuvem — Todos os direitos reservados." no rodapé de cada página.
*   **Interface Simples:** Oferece uma janela intuitiva para que usuários sem conhecimento técnico possam realizar as conversões facilmente.

## 💻 Como Rodar a Aplicação

Para usar o programa, não é necessário instalar o Python ou qualquer uma das bibliotecas. Basta seguir os passos abaixo:

#### Pré-requisito

*   É essencial ter o **Microsoft Word** instalado no seu computador, pois ele é utilizado no processo de conversão.

#### Instruções

1.  Baixe o arquivo executável (`conversor_pdf.exe`) da seção de "Releases" do repositório.
2.  Crie uma nova pasta em seu computador para organizar seus arquivos e mova o `.exe` para dentro dela.
3.  Dê um duplo-clique no arquivo `conversor_pdf.exe` para iniciar o programa.
4.  Na janela da aplicação:
    *   Clique em "Selecionar..." para escolher a imagem da marca d'água.
    *   Clique em "Selecionar..." para escolher os arquivos `.docx` que deseja converter.
    *   Clique no botão verde **"Converter para PDF com Marca d'Água"**.
5.  Aguarde o processo terminar. Os arquivos PDF convertidos serão salvos na mesma pasta onde os arquivos `.docx` originais estão localizados.

---

### Projeto desenvolvido por **Josely Castro**.

[<img src="https://img.shields.io/badge/linkedin-%230077B5.svg?&style=for-the-badge&logo=linkedin&logoColor=white" />](https://www.linkedin.com/in/joselybcastro/) [<img src="https://img.shields.io/badge/github-%23121011.svg?&style=for-the-badge&logo=github&logoColor=white" />](https://github.com/joselyBC)