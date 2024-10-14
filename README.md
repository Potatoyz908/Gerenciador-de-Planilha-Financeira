# Gerenciador de Planilha Financeira

O **Gerenciador de Planilha Financeira** é uma aplicação desenvolvida em Python que permite aos usuários gerenciar e atualizar dados financeiros de forma fácil e eficiente. Através de uma interface gráfica amigável, o usuário pode adicionar, filtrar e organizar informações financeiras em planilhas Excel, facilitando a gestão de centros de custo e fornecedores.

## Funcionalidades

- **Seleção de Arquivo**: O usuário pode selecionar uma planilha Excel existente para adicionar ou atualizar dados financeiros.
- **Adicionar Dados**: Permite a inserção de novos dados financeiros, incluindo data, valor, fornecedor, descrição, centro de custo, observações e outros detalhes.
- **Validação de Dados**: A aplicação valida os dados inseridos, garantindo que todas as informações necessárias estejam corretas antes de serem salvas.
- **Backup Automático**: A aplicação cria backups automáticos do arquivo Excel, permitindo a recuperação de dados em caso de erros ou perda de informações.
- **Atualização Automática de Abas**: Os dados inseridos são automaticamente filtrados e organizados em abas separadas por centro de custo dentro da mesma planilha.
- **Interface Gráfica Intuitiva**: A aplicação conta com uma interface gráfica desenvolvida com a biblioteca Tkinter, facilitando a interação do usuário.

## Tecnologias Utilizadas

- **Python**: Linguagem de programação utilizada para o desenvolvimento da aplicação.
- **Tkinter**: Biblioteca padrão do Python para a criação de interfaces gráficas.
- **Pandas**: Biblioteca para manipulação de dados em formato tabular.
- **Openpyxl**: Biblioteca para leitura e escrita de arquivos Excel.
- **Tkcalendar**: Biblioteca para o componente de seleção de data.

## Como Usar

1. **Instalação das Dependências**:
   Certifique-se de ter o Python instalado em sua máquina. Você pode instalar as dependências necessárias usando o pip:

   ```bash
   pip install pandas openpyxl tkcalendar
2. **Executar a Aplicação: Execute o script Python da aplicação. O usuário será solicitado a selecionar uma planilha Excel para carregar**:
   
   ```bash
   python App.py
3. **Adicionar Dados: Preencha os campos do formulário e clique em "Adicionar Dados" para salvar as informações na planilha. A aplicação irá validar os dados e notificar se houver erros.**
   
4. **Finalizar a Aplicação: Você pode finalizar a aplicação clicando no botão "Finalizar".**
