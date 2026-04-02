import warnings
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
import easyocr
import os

warnings.filterwarnings("ignore", category=UserWarning)

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

    if score >= 85:
        return base_dados[resposta]
    return "Desculpe, não consegui entender sua solicitação... \nPor favor, tente informar a mensagem de erro que aparece na tela."

print("Chatbot iniciado. Digite 'sair' para encerrar a conversa.")

print("Carregando IA para leitura de imagens...")
leitor_imagem = easyocr.Reader(['pt-br'], gpu=False)
print("IA de leitura de imagens carregada com sucesso!")

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

estado_usuarios = {}

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
                
                try:
                    nome_contato = navegador.find_element(By.XPATH, '//*[@id="main"]//header//span[@dir="auto"]').text
                except:
                    nome_contato = "Cliente_Desconhecido"

                baloes_recebidos = navegador.find_elements(By.XPATH, "//div[contains(@class, 'message-in')]")
                
                if baloes_recebidos:
                    ultimo_balao = baloes_recebidos[-1]
                    
                    try:
                        caixa_texto = ultimo_balao.find_element(By.XPATH, ".//div[contains(@class, 'copyable-text')]")
                        data_hora = caixa_texto.get_attribute("data-pre-plain-text")
                    except:
                        data_hora = "[Data/Hora não encontrada] "
                    
                    textos = ultimo_balao.find_elements(By.XPATH, ".//span[@dir='ltr' or contains(@class, 'selectable-text')]")
                    imagens = ultimo_balao.find_elements(By.XPATH, ".//img[contains(@src, 'blob:')]")

                    ultima_mensagem = ""
                    mensagem_limpa = ""

                    if textos:
                        ultima_mensagem = textos[-1].text.lower()
                        mensagem_limpa = ultima_mensagem.strip().replace("!", "").replace("?", "").replace(".", "").replace(",", "")

                    print(f"🕒 {data_hora.strip() if data_hora else ''}")
                    
                    estado_atual = estado_usuarios.get(nome_contato)

                    if estado_atual == "AGUARDANDO_CONFIRMACAO":
                        if mensagem_limpa in ["sim", "s"]:
                            resposta = "Ótimo! Fico feliz em ter ajudado. O seu atendimento foi encerrado. Tenha um excelente dia!"
                            estado_usuarios.pop(nome_contato, None)
                            
                        elif mensagem_limpa in ["nao", "não", "n", "nao resolveu", "não resolveu"]:
                            resposta = "Certo, entendi. Estou encaminhando o seu caso para um de nossos atendentes. Por favor, aguarde um momento."
                            estado_usuarios.pop(nome_contato, None)
                            
                        else:
                            resposta = "Por favor, responda apenas com *Sim* ou *Não*. A solução que enviei anteriormente resolveu o seu problema?"

                    else:
                        if imagens:
                            print("📷 Imagem detectada. O robô está processando...")
                            try:
                                imagem_elemento = imagens[-1]
                                caminho_imagem = "imagem_recebida.png"

                                imagem_elemento.screenshot(caminho_imagem)

                                resultado = leitor_imagem.readtext(caminho_imagem, detail=0)
                                texto_completo = ' '.join(resultado).lower()
                                
                                # LÓGICA NOVA: Prioriza o texto DEPOIS da palavra "rejeição"
                                if "rejeição" in texto_completo:
                                    texto_extraido = texto_completo.split("rejeição", 1)[1]
                                elif "rejeicao" in texto_completo:
                                    texto_extraido = texto_completo.split("rejeicao", 1)[1]
                                else:
                                    texto_extraido = texto_completo
                                
                                # Remove possíveis dois pontos, traços ou espaços no início da frase cortada
                                texto_extraido = texto_extraido.strip(' :-=>')

                                print(f"👀 O robô filtrou a imagem e buscará por: '{texto_extraido}'")

                                if len(texto_extraido.strip()) > 4:
                                    resposta = buscar_resposta(texto_extraido)
                                    
                                    if "Desculpe, não consegui entender sua solicitação" in resposta:
                                        resposta = "Desculpe, não encontrei a solução para o erro mostrado na imagem. Por favor, tente enviar uma foto mais nítida ou digite o erro manualmente."
                                    else:
                                        resposta += "\n\nEssa solução resolveu o seu problema? (Responda *Sim* ou *Não*)"
                                        estado_usuarios[nome_contato] = "AGUARDANDO_CONFIRMACAO"
                                        
                                else:
                                    resposta = "Desculpe, mas a imagem parece estar ilegível ou sem texto de erro. Por favor, tente enviar uma imagem mais clara ou informe o erro manualmente."
                                
                                if os.path.exists(caminho_imagem):
                                    os.remove(caminho_imagem)

                            except Exception as e:
                                print(f"⚠️ Ocorreu um erro ao processar a imagem: {e}")
                                resposta = "Desculpe, mas ocorreu um erro interno ao processar a imagem. Por favor, digite o erro manualmente."

                        elif mensagem_limpa:
                            print(f"👀 O robô leu a mensagem: '{mensagem_limpa}'")
                            saudacoes = ["oi", "olá", "ola", "bom dia", "boa tarde", "boa noite", "preciso de ajuda", "menu", "deu erro"]

                            if mensagem_limpa in saudacoes:
                                    resposta = ("Olá! Sou o assistente virtual de suporte Insoft.\n\n"
                                            "Para que eu possa te ajudar, por favor, informe qual a sua solicitação:\n"
                                            "A - NFe\n"
                                            "B - Vendas\n"
                                            "C - Outros\n"
                                            "E - Sair")

                            elif mensagem_limpa in ["a", "nfe"]:
                                resposta = "Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre NFe."

                            elif mensagem_limpa in ["b", "vendas"]:
                                resposta = "Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre vendas."

                            elif mensagem_limpa in ["c", "outros"]:
                                resposta = "Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre outros assuntos."

                            elif mensagem_limpa in ["e", "sair"]:
                                resposta = "Obrigado por entrar em contato com o suporte Insoft. Se precisar de mais ajuda, é só chamar. Tenha um ótimo dia!"
                            
                            else:
                                resposta = buscar_resposta(ultima_mensagem)
                                
                                if "Desculpe" not in resposta:
                                    resposta += "\n\nEssa solução resolveu o seu problema? (Responda *Sim* ou *Não*)"
                                    estado_usuarios[nome_contato] = "AGUARDANDO_CONFIRMACAO"

                        else:
                            print("❌ Formato não suportado")
                            resposta = "Desculpe, mas o tipo de mensagem que você enviou não é suportado pelo nosso robô. Por favor, envie uma mensagem de texto ou uma imagem clara do erro que você está enfrentando."

                    
                    caixa_texto_envio = navegador.find_element(By.XPATH, '//*[@id="main"]//footer//div[@contenteditable="true"]')

                    for linha in resposta.split('\n'):
                        linha_sem_emoji = "".join(c for c in linha if ord(c) <= 0xFFFF)
                        caixa_texto_envio.send_keys(linha_sem_emoji)
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
