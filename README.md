# Automação de Saldo Cliente - WeChat 🚀

Este projeto automatiza o processamento de saldos de clientes e o envio de comprovantes/saldos através do WeChat Desktop. Ele extrai informações diretamente de planilhas Excel, gera imagens de alta qualidade e realiza a colagem automática nos chats dos clientes.

## 🛠️ Funcionalidades

- **Atualização de Diário:** Processa a planilha de saldos do dia.
- **Extração de Imagens:** Gera imagens nítidas (High DPI) dos comprovantes e do saldo atual do cliente.
- **Mapeamento de Grupos:** Associa automaticamente cada cliente ao seu respectivo grupo no WeChat.
- **Automação de Colagem:** Localiza os chats e cola as imagens como rascunho, permitindo conferência final antes do envio humano.

## 📂 Estrutura do Projeto

O fluxo é dividido em passos numerados para facilitar a execução:

1.  **`1 - Atualizar Diario Saldo.bat`**: Sincroniza e prepara os dados na planilha Excel.
2.  **`2 - Preparar Envio WeChat.bat`**: 
    - Extrai as imagens do Excel.
    - Gera um relatório de conferência em HTML (`Conferir_Imagens.html`).
    - Cria uma fila de envio em CSV.
3.  **`3 - Colar WeChat Rascunhos.bat`**: Abre o WeChat, pesquisa os grupos e cola as imagens extraídas automaticamente.

## ⚙️ Configuração

- **Mapeamento de Clientes:** Edite o arquivo `WechatClienteMap.json` para associar o nome do cliente na planilha ao nome exato do grupo no WeChat.
- **Caminho da Planilha:** O script solicitará a localização da planilha `.xlsx` ou `.xlsm` na primeira execução ou você pode configurar os caminhos padrão nos arquivos `.ps1`.

## 📋 Requisitos

- **Windows 10/11**
- **Microsoft Excel** instalado.
- **WeChat Desktop** aberto e logado.
- **PowerShell** habilitado para execução de scripts.

## 💎 Diferenciais Técnicos

- **Qualidade de Imagem:** Utiliza captura vetorial (`xlPicture`) e escala de 2.0x para garantir que as tabelas fiquem perfeitamente legíveis mesmo no celular.
- **Segurança:** A automação cola as imagens como **rascunho**, garantindo que nenhum dado seja enviado sem uma revisão visual prévia.

---
Desenvolvido para otimizar o fluxo de atendimento e prestação de contas.
