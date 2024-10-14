# 📊 Gerenciador de Planilha Financeira

O **Gerenciador de Planilha Financeira** é uma aplicação desenvolvida em Python que permite aos usuários gerenciar e atualizar dados financeiros de forma fácil e eficiente. Através de uma interface gráfica amigável, o usuário pode adicionar, filtrar e organizar informações financeiras em planilhas Excel, facilitando a gestão de centros de custo e fornecedores.

## 🚀Funcionalidades

- **Seleção de Arquivo**: O usuário pode selecionar uma planilha Excel existente para adicionar ou atualizar dados financeiros.
- **Adicionar Dados**: Permite a inserção de novos dados financeiros, incluindo data, valor, fornecedor, descrição, centro de custo, observações e outros detalhes.
- **Validação de Dados**: A aplicação valida os dados inseridos, garantindo que todas as informações necessárias estejam corretas antes de serem salvas.
- **Backup Automático**: A aplicação cria backups automáticos do arquivo Excel, permitindo a recuperação de dados em caso de erros ou perda de informações.
- **Atualização Automática de Abas**: Os dados inseridos são automaticamente filtrados e organizados em abas separadas por centro de custo dentro da mesma planilha.
- **Interface Gráfica Intuitiva**: A aplicação conta com uma interface gráfica desenvolvida com a biblioteca Tkinter, facilitando a interação do usuário.

## 🛠 Tecnologias Utilizadas

- **Python**: Linguagem de programação utilizada para o desenvolvimento da aplicação.
- **Tkinter**: Biblioteca padrão do Python para a criação de interfaces gráficas.
- **Pandas**: Biblioteca para manipulação de dados em formato tabular.
- **Openpyxl**: Biblioteca para leitura e escrita de arquivos Excel.
- **Tkcalendar**: Biblioteca para o componente de seleção de data.

## 📥 Instalação

1. **Certifique-se de ter o Python 3.8 ou superior instalado em sua máquina.** Você pode baixar a versão mais recente do Python no site oficial:

   [![Baixar Python](https://img.shields.io/badge/Download_Python-blue)](https://www.python.org/downloads/)

2. **Instale as dependências do projeto.** Navegue até o diretório onde o projeto está localizado e execute o seguinte comando no terminal:

   ```bash
   pip install -r requirements.txt
   
3. **O projeto também inclui um executável.** O arquivo executável é chamado `App.exe`. Você pode executá-lo diretamente clicando duas vezes sobre o arquivo.

## ☕ Como executar o projeto (WINDOWS)

1. **Execute o arquivo `App.exe` diretamente.** Navegue até o diretório onde o arquivo está localizado e clique duas vezes sobre ele para iniciar o aplicativo.

2. **Outra forma é executar o arquivo pelo terminal:**  
   Abra o terminal (cmd) e entre no diretório em que o arquivo `App.exe` se encontra:
   ```bash
   cd C:\Caminho\Para\O\Diretório
   
**Substitua `C:\Caminho\Para\O\Diretório` pelo caminho real onde o arquivo `App.exe` está.**

## ☕ Como executar o projeto (LINUX-DEBIAN)

1. **Certifique-se de ter o Python 3.8 ou superior instalado.** Se necessário, você pode instalar o Python usando o seguinte comando no terminal:
   ```bash
   sudo apt update
   sudo apt install python3
   
2. **Instale as dependências do projeto. Navegue até o diretório onde o projeto está localizado e execute:**
   ```bash
   pip install -r requirements.

3. **Para executar o projeto:
Entre no diretório onde o arquivo executável está localizado:**
   ```bash
   cd /caminho/para/o/diretório

4. **Substitua `/caminho/para/o/diretório` pelo caminho real onde o arquivo `App.exe` está.
Agora use o comando para executar o aplicativo:**
   ```bash
   ./App.exe
5. **Certifique-se de que o arquivo `App.exe` tenha permissões de execução. Se necessário, ajuste as permissões com:**
   ```bash
   chmod +x App.exe


## 🐍 Como executar o ARQUIVO.PY

1. **Instalação das Dependências**:
   Certifique-se de ter o Python instalado em sua máquina. Você pode instalar as dependências necessárias usando o pip:

   ```bash
   pip install pandas openpyxl tkcalendar
   
2. **Executar a Aplicação: Execute o script Python da aplicação. O usuário será solicitado a selecionar uma planilha Excel para carregar**:
   
   ```bash
   python App.py
   
3. **Adicionar Dados: Preencha os campos do formulário e clique em "Adicionar Dados" para salvar as informações na planilha. A aplicação irá validar os dados e notificar se houver erros.**
   
4. **Finalizar a Aplicação: Você pode finalizar a aplicação clicando no botão "Finalizar".**

## 💻 Exemplos de execução

### Imagens de funcionamento do programa:

  ![](https://github.com/Potatoyz908/Gerenciador-de-Planilha-Financeira/blob/main/imgs/Captura1.png)  ![](https://github.com/Potatoyz908/Gerenciador-de-Planilha-Financeira/blob/main/imgs/Captura2.png)  ![](https://github.com/Potatoyz908/Gerenciador-de-Planilha-Financeira/blob/main/imgs/Captura3.png)
 
