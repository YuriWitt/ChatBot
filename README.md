#  Assistente Virtual de Suporte WhatsApp

Projeto de um chatbot automatizado em Python que opera através do WhatsApp Web. Ele foi desenvolvido para realizar a triagem inicial de clientes, identificar erros descritos em texto ou capturados em imagens (screenshots), buscar soluções em uma base de conhecimento (planilha Excel) e transferir o atendimento para um atendente humano quando necessário.



* **Atendimento Automatizado:** Responde automaticamente a novas mensagens no WhatsApp Web.
* **Validação de Horário:** Verifica se o contato está sendo feito dentro do horário de expediente (Seg a Sex).
* **Leitura de Imagens (OCR):** Utiliza Visão Computacional para ler capturas de tela enviadas pelo usuário, identificando mensagens de rejeição da SEFAZ ou outros erros.
* **Busca Inteligente (Fuzzy Matching):** Analisa a mensagem de erro do cliente e busca a resposta mais adequada na `Base_De_Conhecimento.xlsx`, mesmo que haja pequenos erros de digitação.
* **Máquina de Estados:** Gerencia o fluxo da conversa (Solicitação de CNPJ/Nome -> Menu de Opções -> Suporte -> Confirmação de Resolução -> Avaliação).
* **Transbordo Humano:** Interrompe a automação de forma inteligente caso o cliente não tenha seu problema resolvido ou se um atendente humano intervir na conversa.
* **Pesquisa de Satisfação:** Coleta uma nota de 1 a 5 ao finalizar o atendimento.

## Tecnologias Utilizadas

* **Python 3.13** - Linguagem principal do projeto.
* **Selenium & WebDriver Manager** - Para automação e controle do Google Chrome no WhatsApp Web.
* **EasyOCR & OpenCV (cv2)** - Para processamento de imagens e extração de texto (Optical Character Recognition).
* **Pandas** - Para leitura e manipulação da base de conhecimento em Excel.
* **RapidFuzz** - Para comparação de strings e busca de similaridade de texto.
* **PyInstaller** - Para conversão do script em um arquivo executável (`.exe`).

## Pré-requisitos para Execução

Para que o executável ou o script funcione corretamente na máquina destino, é necessário garantir que:

1. O **Google Chrome** esteja instalado no computador.
2. O computador tenha acesso ao disco de rede mapeado como `R:`, especificamente ao caminho: `R:\Sistemas\Manuais\Base_De_Conhecimento.xlsx`.

## Como Executar (Versão Compilada)

1. Navegue até a pasta onde o arquivo `ChatBot.exe` está salvo.
2. Dê um duplo clique no arquivo. *(Na primeira execução de cada dia, pode levar alguns segundos para o terminal aparecer, pois o sistema extrai as bibliotecas pesadas de IA e Visão Computacional).*
3. Uma janela do Google Chrome se abrirá automaticamente na página do WhatsApp Web.
4. Escaneie o **QR Code** com o celular da empresa.
5. Pronto! O robô criará uma pasta chamada `sessao_whatsapp` no diretório do usuário do Windows para salvar o login. Nas próximas vezes, não será necessário escanear o QR Code novamente.

## Lógica de Atendimento (Fluxo)

1. **`AGUARDANDO_EMPRESA_CNPJ`**: Pede o nome da empresa/CNPJ.
2. **`AGUARDANDO_NOME`**: Pede o nome da pessoa que está falando.
3. **`AGUARDANDO_MENU`**: Oferece opções (A - NFe, B - Vendas, C - Outros, E - Encerrar).
4. **`EM_SUPORTE`**: Analisa textos ou imagens de erro enviadas pelo cliente.
5. **`AGUARDANDO_CONFIRMACAO`**: Confirma se a solução enviada pelo robô resolveu o problema (Sim/Não).
6. **`ATENDIMENTO_HUMANO`**: Estado onde o robô ignora o cliente para que um atendente real possa ajudar.
7. **`AGUARDANDO_AVALIACAO`**: Pede a nota final (1 a 5) antes de encerrar o chamado.

## Observações e Limitações

* **Não minimize ou feche o navegador Chrome** controlado pelo robô, pois o Selenium precisa que os elementos da página estejam visíveis ou carregados para interagir com eles.
* Se a estrutura HTML do WhatsApp Web for atualizada pela Meta (Facebook), os seletores XPath do script podem precisar de manutenção.****
