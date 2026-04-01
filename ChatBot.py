from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from rapidfuzz import process
import time
import easyocr # Importando a Inteligência Artificial
import os

def buscar_resposta(mensagem):
    try:
        df = pd.read_excel(r'R:\Sistemas\Manuais\Base_De_Conhecimento.xlsx')
        df.columns = df.columns.str.strip()
        base_dados = dict(zip(df['Rejeição'].astype(str), df['Solução & Informações Adicionais'].astype(str)))
    except Exception as e:
        return f"Erro ao ler a base de dados: {str(e)}" 
    
    if not base_dados:
        return "Base de dados vazia ou não encontrada."
    
    preguntas = list(base_dados.keys())
    resposta, score, _ = process.extractOne(mensagem, preguntas)

    if score >= 80:
        return base_dados[resposta]
    return "Desculpe, não consegui entender sua solicitação... \nPor favor, tente informar a mensagem de erro que aparece na tela."

print("Chatbot iniciado. Digite 'sair' para encerrar a conversa.")

# Carregando a Inteligência Artificial antes de abrir o navegador (pode levar alguns segundos)
print("🧠 Carregando o motor de IA para leitura de imagens...")
leitor_ia = easyocr.Reader(['pt']) # 'pt' para português
print("✅ Motor de IA carregado com sucesso!")

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.get("https://web.whatsapp.com/")

print("Aguardando o QR code ser escaneado...")

try:
    WebDriverWait(navegador, 120).until(EC.presence_of_element_located((By.ID, "side")))
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
            try:
                bolinha.click()
                time.sleep(2)
                
                grupo_check = navegador.find_elements(By.XPATH, '//*[@id="main"]//div[contains(@data-id, "@g.us")]')
                if grupo_check:
                    print("🚫 Mensagem de grupo detectada. Ignorando e fechando...")
                    acoes = ActionChains(navegador)
                    acoes.send_keys(Keys.ESCAPE).perform()
                    time.sleep(1)
                    continue
                
                baloes_recebidos = navegador.find_elements(By.XPATH, "//div[contains(@class, 'message-in')]")
                
                if baloes_recebidos:
                    ultimo_balao = baloes_recebidos[-1]
                    
                    try:
                        caixa_texto = ultimo_balao.find_element(By.XPATH, ".//div[contains(@class, 'copyable-text')]")
                        data_hora = caixa_texto.get_attribute("data-pre-plain-text")
                    except:
                        data_hora = "[Data/Hora não encontrada] "
                    
                    # 1. Procurando Textos ou Imagens no último balão
                    textos = ultimo_balao.find_elements(By.XPATH, ".//span[@dir='ltr' or contains(@class, 'selectable-text')]")
                    imagens = ultimo_balao.find_elements(By.XPATH, ".//img[contains(@src, 'blob')]")
                    
                    ultima_mensagem = ""
                    if textos:
                        ultima_mensagem = textos[-1].text.lower()
                    
                    mensagem_limpa = ultima_mensagem.strip().replace("!", "").replace("?", "").replace(".", "").replace(",", "")

                    print(f"🕒 {data_hora.strip() if data_hora else ''}")

                    # 2. Lógica de Decisão (Texto x Imagem)
                    if mensagem_limpa:
                        print(f"👀 O robô leu a mensagem de texto: '{mensagem_limpa}'")
                        saudacoes = ["oi", "olá", "ola", "bom dia", "boa tarde", "boa noite", "preciso de ajuda", "menu", "deu erro"]
                        
                        if mensagem_limpa in saudacoes:
                                resposta = ("Olá! Sou o assistente virtual de suporte Insoft.\n\n"
                                        "Para que eu possa te ajudar, por favor, informe qual a sua solicitação:\n"
                                        "A - NFe\n"
                                        "B - Vendas\n"
                                        "C - Outros\n"
                                        "E - Sair")
                        elif mensagem_limpa in ["a", "nfe"]:
                            resposta = "Entendi que você precisa de ajuda com NFe. Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre NFe."
                        elif mensagem_limpa in ["b", "vendas"]:
                            resposta = "Entendi que você precisa de ajuda com vendas. Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre vendas."
                        elif mensagem_limpa in ["c", "outros"]:
                            resposta = "Entendi que você precisa de ajuda com outros assuntos. Por favor, informe o erro que aparece na tela ou a dúvida que você tem."
                        elif mensagem_limpa in ["e", "sair"]:
                            resposta = "Obrigado por entrar em contato com o suporte Insoft. Se precisar de mais ajuda, é só chamar. Tenha um ótimo dia!"
                        else:
                            resposta = buscar_resposta(ultima_mensagem)

                    elif imagens:
                        print("📸 Imagem detectada! Iniciando a IA de leitura (OCR)...")
                        try:
                            imagem_elemento = imagens[-1]
                            caminho_imagem = "print_cliente.png"
                            
                            # Tira print do balão da imagem
                            imagem_elemento.screenshot(caminho_imagem)
                            
                            # A IA lê a imagem e junta tudo em uma string
                            resultados_ia = leitor_ia.readtext(caminho_imagem, detail=0)
                            texto_extraido = " ".join(resultados_ia).lower()
                            
                            print(f"🧠 IA extraiu o texto da imagem: '{texto_extraido}'")
                            
                            if len(texto_extraido) > 4:
                                resposta = buscar_resposta(texto_extraido)
                                if "Desculpe, não consegui entender" in resposta:
                                    resposta = ("Li a sua imagem, mas não encontrei a solução exata para esse erro na minha base de dados. 😕\n\n"
                                                "Você poderia digitar apenas o código da rejeição ou a mensagem principal?")
                            else:
                                resposta = ("Vi que você mandou uma imagem, mas não consegui identificar nenhum texto legível nela. 🧐\n\n"
                                            "Por favor, digite a mensagem de erro que aparece na tela.")
                                            
                            # Opcional: apaga a imagem do PC depois de ler para não acumular lixo
                            if os.path.exists(caminho_imagem):
                                os.remove(caminho_imagem)
                                
                        except Exception as e:
                            print(f"⚠️ Erro ao tentar ler a imagem com IA: {e}")
                            resposta = "Houve uma falha ao tentar ler a sua imagem. Por favor, digite o erro para que eu possa ajudar!"
                    
                    else:
                        print("📎 O robô detectou um formato não suportado.")
                        resposta = ("Recebi sua mensagem, mas eu só consigo entender textos digitados ou prints de tela. ⌨️📸\n\n"
                                    "Por favor, digite a sua solicitação ou a mensagem de erro.")

                    # 3. Enviando a resposta
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
                    acoes = ActionChains(navegador)
                    acoes.send_keys(Keys.ESCAPE).perform()

            except StaleElementReferenceException:
                print("⚠️ A página do WhatsApp atualizou enquanto o robô lia. Recalculando...")
                acoes = ActionChains(navegador)
                acoes.send_keys(Keys.ESCAPE).perform()
                continue

    except TimeoutException:
        pass
    except Exception as e:
        print(f"⚠️ Ocorreu um erro na automação: {e}")
        time.sleep(2)
          
    time.sleep(3)
