from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from rapidfuzz import process
import time

def buscar_resposta(mensagem):
    try:
        df = pd.read_excel(r'R:\Sistemas\Manuais\BC.xlsx')
        df.columns = df.columns.str.strip()
        base_dados = dict(zip(df['Rejeição'].astype(str), df['Solução & Informações Adicionais'].astype(str)))
    except Exception as e:
        return f"Erro ao ler a base de dados: {str(e)}" 
    
    if not base_dados:
        return "Base de dados vazia ou não encontrada."
    
    preguntas = list(base_dados.keys())
    resposta, score, _ = process.extractOne(mensagem, preguntas)

    if score >= 70:
        return base_dados[resposta]
    return "Desculpe, não consegui encontrar uma resposta adequada para a sua pergunta. " \
        "Por favor, tente informar o erro que aparece na tela."

print("Chatbot iniciado. Digite 'sair' para encerrar a conversa.")
chrome_options = Options()

chrome_options.add_experimental_option("detach", True)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.get("https://web.whatsapp.com/")

print("Aguardando o QR code ser escaneado...")

try:
    WebDriverWait(navegador, 120).until( EC.presence_of_element_located((By.ID, "side")))
    print("QR code escaneado. Chatbot pronto para uso.")
except TimeoutException:
    print("Tempo de espera esgotado para escanear o QR code. Por favor, reinicie o chatbot e tente novamente.")
    navegador.quit()
    exit()


while True:
    try:
        mensagens_nao_lidas = navegador.find_elements(By.XPATH, "//span[contains(@aria-label, 'lida')]")

        if len(mensagens_nao_lidas) > 0:
            print(f"🔔 Você tem {len(mensagens_nao_lidas)} mensagens não lidas. O robô está lendo...")

        for bolinha in mensagens_nao_lidas:
            
            bolinha.click()
            time.sleep(3)
            
            # Pega o último balão de mensagem da conversa que acabou de abrir
            baloes_recebidos = navegador.find_elements(By.XPATH, "//div[contains(@class, 'message-in')]")
            
            if baloes_recebidos:
                # Mantemos como "Elemento" do Selenium para poder investigar dentro dele
                ultimo_balao = baloes_recebidos[-1]
                
                # ==========================================
                # DETECTOR DE GRUPOS DEFINITIVO (@g.us)
                # ==========================================
                data_id = ultimo_balao.get_attribute("data-id")
                if data_id and "@g.us" in data_id:
                    print("🚫 Mensagem de grupo detectada. Ignorando e fechando...")
                    acoes = ActionChains(navegador)
                    acoes.send_keys(Keys.ESCAPE).perform()
                    continue # Aborta tudo e pula para o próximo cliente
                # ==========================================
                
                # 1. CAPTURANDO DATA E HORA
                try:
                    caixa_texto = ultimo_balao.find_element(By.XPATH, ".//div[contains(@class, 'copyable-text')]")
                    data_hora = caixa_texto.get_attribute("data-pre-plain-text")
                except:
                    data_hora = "[Data/Hora não encontrada] "
                
                # 2. CAPTURANDO O TEXTO DA MENSAGEM
                textos = ultimo_balao.find_elements(By.XPATH, ".//span[@dir='ltr' or contains(@class, 'selectable-text')]")
                
                if textos:
                    ultima_mensagem = textos[-1].text.lower()
                    mensagem_limpa = ultima_mensagem.strip().replace("!", "").replace("?", "").replace(".", "").replace(",", "")

                    print(f"🕒 {data_hora.strip() if data_hora else ''}")
                    print(f"👀 O robô leu a mensagem: '{mensagem_limpa}'")

                    saudacoes = ["oi", "olá", "ola", "bom dia", "boa tarde", "boa noite", "preciso de ajuda", "menu", "deu erro"]
                    
                    if mensagem_limpa in saudacoes:
                            resposta = ("Olá! Sou o assistente virtual de suporte Insoft.\n\n"
                                    "Para que eu possa te ajudar, por favor, informe qual a sua solicitação:\n"
                                    "A - NFe\n"
                                    "B - Vendas\n"
                                    "C - Outros\n"
                                    "E - Sair")

                    elif mensagem_limpa == "a" or mensagem_limpa == "nfe":
                        resposta = "Entendi que você precisa de ajuda com NFe. Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre NFe."

                    elif mensagem_limpa == "b" or mensagem_limpa == "vendas":
                        resposta = "Entendi que você precisa de ajuda com vendas. Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre vendas."

                    elif mensagem_limpa == "c" or mensagem_limpa == "outros":
                        resposta = "Entendi que você precisa de ajuda com outros assuntos. Por favor, informe o erro que aparece na tela ou a dúvida que você tem."

                    elif mensagem_limpa == "e" or mensagem_limpa == "sair":
                        resposta = "Obrigado por entrar em contato com o suporte Insoft. Se precisar de mais ajuda, é só chamar. Tenha um ótimo dia!"
                    
                    else:
                        resposta = buscar_resposta(ultima_mensagem)

                    caixa_texto_envio = navegador.find_element(By.XPATH, '//*[@id="main"]//footer//div[@contenteditable="true"]')

                    for linha in resposta.split('\n'):
                        caixa_texto_envio.send_keys(linha)
                        caixa_texto_envio.send_keys(Keys.SHIFT , Keys.ENTER)

                    caixa_texto_envio.send_keys(Keys.ENTER)

                    print(f"✅ Resposta enviada com sucesso!")
                    time.sleep(2)

                    acoes = ActionChains(navegador)
                    acoes.send_keys(Keys.ESCAPE).perform()

            else:
                print("❌ Abri a conversa, mas não consegui ler o texto da mensagem.")

    except TimeoutException:
        pass
    except Exception as e:
        print(f"⚠️ Ocorreu um erro na automação: {e}")
        time.sleep(2)
          
    time.sleep(3)
