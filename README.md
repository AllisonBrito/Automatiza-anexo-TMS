# Automação de Upload de Anexos para TMS

Este é um script em Python que automatiza o processo de upload de anexos para um sistema de gerenciamento de transporte (TMS). O script utiliza a biblioteca Selenium para realizar a automação de tarefas no navegador.

## Pré-requisitos

Python 3.x
Bibliotecas Python: 
selenium, pandas, webdriver_manager
Navegador Google Chrome

## Como Usar

Clone ou faça o download deste repositório para o seu ambiente local.

Certifique-se de ter todas as bibliotecas Python necessárias instaladas. Você pode instalar as dependências com o seguinte comando:

pip install selenium pandas webdriver_manager

Crie uma nova pasta com o nome anexo_tms no diretório raiz do disco C:.

Faça o download do arquivo Excel necessário do repositório e inclua-o na pasta recém-criada anexo_tms.

Edite o script para fornecer os detalhes de login do TMS e os caminhos para o arquivo Excel que contém os dados necessários para o upload dos anexos.

Execute o script Python. Ele abrirá uma instância do navegador Chrome e realizará o login no TMS. Em seguida, ele percorrerá os vouchers especificados no arquivo Excel e fará o upload dos anexos conforme necessário.

## Funcionamento

Este script realiza as seguintes etapas:

Lê os dados de login, URL e caminho do anexo de um arquivo Excel.
Configura o WebDriver e abre o navegador Chrome.
Realiza o login no TMS utilizando os dados fornecidos.
Navega até a tela de vouchers no TMS.
Insere os anexos para cada voucher especificado no arquivo Excel.
Finaliza o processo e fecha o navegador.
