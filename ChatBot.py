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
from rapidfuzz import process, fuzz
import time
import easyocr
import os
import cv2 
import re   

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
    
    resposta, score, _ = process.extractOne(mensagem, preguntas, scorer=fuzz.token_set_ratio)

    print(f"[DEBUG DA IA] Match encontrado: '{resposta}' | Pontuação: {score}")

    if score >= 60:
        return base_dados[resposta]
    return "Desculpe, não consegui entender sua solicitação... \nPor favor, tente informar a mensagem de erro que aparece na tela."

print("Chatbot iniciado. Digite 'sair' para encerrar a conversa.")

print("Carregando a leitura de imagens...")
leitor_imagem = easyocr.Reader(['pt'], gpu=False)
print("Leitura de imagens carregada com sucesso!")

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
tentativas_falhas = {}

while True:
    try:
        mensagens_nao_lidas = navegador.find_elements(By.XPATH, "//*[@id='pane-side']//span[contains(@aria-label, 'lida')]")

        if len(mensagens_nao_lidas) > 0:
            print(f"Você tem {len(mensagens_nao_lidas)} mensagens não lidas. O robô está lendo...")

        for bolinha in mensagens_nao_lidas:
            try:
                bolinha.click()
                time.sleep(2)
                
                grupo_check = navegador.find_elements(By.XPATH, '//*[@id="main"]//div[contains(@data-id, "@g.us")]')
                if grupo_check:
                    print("Mensagem de grupo detectada. Ignorando e fechando...")
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
                        texto_bruto = caixa_texto.text
                    except:
                        data_hora = "[Data/Hora não encontrada] "
                        texto_bruto = ""
                    
                    imagens = ultimo_balao.find_elements(By.XPATH, ".//img[contains(@src, 'blob:')]")

                    ultima_mensagem = ""
                    mensagem_limpa = ""

                    if texto_bruto.strip() != "":
                        ultima_mensagem = texto_bruto.lower().strip()
                        mensagem_limpa = ultima_mensagem.replace("!", "").replace("?", "").replace(".", "").replace(",", "")
                            
                    print(f"{data_hora.strip() if data_hora else ''}")
                    
                    estado_atual = estado_usuarios.get(nome_contato)

                    if estado_atual == "ATENDIMENTO_HUMANO":
                        baloes_enviados = navegador.find_elements(By.XPATH, "//div[contains(@class, 'message-out')]")
                        atendimento_finalizado = False
                        
                        if baloes_enviados:
                            ultimo_balao_enviado = baloes_enviados[-1]
                            try:
                                texto_enviado = ultimo_balao_enviado.find_element(By.XPATH, ".//div[contains(@class, 'copyable-text')]").text.lower()
                                if "agradecemos pelo contato" in texto_enviado and "seu atendimento foi finalizado" in texto_enviado:
                                    atendimento_finalizado = True
                            except:
                                pass
                                
                        if atendimento_finalizado:
                            print(f"Atendimento de '{nome_contato}' finalizado pelo humano. Retomando o controle e enviando avaliação...")
                            resposta = "Para nos ajudar a melhorar, como você avalia o meu atendimento de *1 a 5*? (Sendo 1 Ruim e 5 Excelente)"
                            estado_usuarios[nome_contato] = "AGUARDANDO_AVALIACAO"
                        else:
                            print(f"O contato '{nome_contato}' continua conversando com um atendente humano. O robô vai ignorar.")
                            acoes = ActionChains(navegador)
                            acoes.send_keys(Keys.ESCAPE).perform()
                            continue
                    
                    elif estado_atual == "AGUARDANDO_CONFIRMACAO":
                        if mensagem_limpa in ["sim", "s", "sim resolveu", "resolvido", "resolvi"]:
                            resposta = "Ótimo! Fico feliz em ter ajudado.\n\nPara nos ajudar a melhorar, como você avalia o meu atendimento de *1 a 5*? (Sendo 1 Ruim e 5 Excelente)"
                            estado_usuarios[nome_contato] = "AGUARDANDO_AVALIACAO"
                            
                        elif mensagem_limpa in ["nao", "não", "n", "nao resolveu", "não resolveu"]:
                            resposta = "Certo, entendi. Estou encaminhando o seu caso para um de nossos atendentes. Por favor, aguarde um momento."
                            estado_usuarios[nome_contato] = "ATENDIMENTO_HUMANO"
                            
                        else:
                            resposta = "Por favor, responda apenas com *Sim* ou *Não*. A solução que enviei anteriormente resolveu o seu problema?"

                    elif estado_atual == "AGUARDANDO_AVALIACAO":
                        match_nota = re.search(r"[1-5]", mensagem_limpa)
                        
                        if match_nota:
                            nota = match_nota.group(0)
                            resposta = "Muito obrigado pela sua avaliação! Seu feedback é fundamental para nós. O seu atendimento foi encerrado. Tenha um excelente dia!"
                            print(f"AVALIAÇÃO: O cliente '{nome_contato}' avaliou o atendimento com nota {nota}!")
                            estado_usuarios.pop(nome_contato, None)
                        else:
                            print(f"O robô não encontrou um número de 1 a 5. O texto lido foi: '{mensagem_limpa}'")
                            resposta = "Por favor, digite apenas um número de *1 a 5* para avaliar o atendimento."

                    else:
                        if imagens:
                            print("📷 Imagem detectada. Ampliando para ler com qualidade máxima...")
                            try:
                                imagem_elemento = imagens[-1]
                                caminho_imagem = "imagem_recebida.png"

                                imagem_elemento.click()
                                time.sleep(2)

                                navegador.save_screenshot(caminho_imagem)

                                acoes = ActionChains(navegador)
                                acoes.send_keys(Keys.ESCAPE).perform()
                                time.sleep(1)

                                img = cv2.imread(caminho_imagem)
                                cinza = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                                
                                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                                img_clahe = clahe.apply(cinza)
                                
                                ampliada_clahe = cv2.resize(img_clahe, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
                                
                                cv2.imwrite(caminho_imagem, ampliada_clahe)

                                resultado = leitor_imagem.readtext(caminho_imagem, detail=0)
                                texto_completo = ' '.join(resultado).lower()
                                
                                if 'sefaz' in texto_completo:
                                    texto_extraido = 'sefaz' + texto_completo.split('sefaz', 1)[1]
                                else:
                                    partes = re.split(r'(?:rejei[cç][aã]o|rejei.*?|rjcicoo)[:\s]*', texto_completo)
                                    texto_extraido = partes[1] if len(partes) > 1 else texto_completo

                                texto_extraido = re.sub(r'\d{2,}', '', texto_extraido)
                                texto_extraido = re.sub(r'[^\w\s]', ' ', texto_extraido)
                                texto_extraido = re.sub(r'\s+', ' ', texto_extraido).strip()

                                print(f"O robô filtrou a imagem e buscará por: '{texto_extraido}'")

                                if len(texto_extraido.strip()) > 4:
                                    resposta = buscar_resposta(texto_extraido)
                                    
                                    if "Desculpe, não consegui entender sua solicitação" in resposta:
                                        falhas = tentativas_falhas.get(nome_contato, 0) + 1
                                        if falhas >= 2:
                                            resposta = "Parece que não estou conseguindo encontrar a solução para o seu problema.\n\nEstou encaminhando o seu caso para um de nossos atendentes. Por favor, aguarde um momento."
                                            estado_usuarios[nome_contato] = "ATENDIMENTO_HUMANO"
                                            tentativas_falhas.pop(nome_contato, None)
                                        else:
                                            resposta = "Desculpe, não encontrei a solução para o erro mostrado na imagem. Por favor, tente enviar uma foto mais nítida ou digite o erro manualmente."
                                            tentativas_falhas[nome_contato] = falhas
                                    else:
                                        tentativas_falhas[nome_contato] = 0
                                        resposta += "\n\nEssa solução resolveu o seu problema? (Responda *Sim* ou *Não*)"
                                        estado_usuarios[nome_contato] = "AGUARDANDO_CONFIRMACAO"
                                        
                                else:
                                    falhas = tentativas_falhas.get(nome_contato, 0) + 1
                                    if falhas >= 2:
                                        resposta = "Parece que não estou conseguindo ler a imagem do seu problema.\n\nEstou encaminhando o seu caso para um de nossos atendentes. Por favor, aguarde um momento."
                                        estado_usuarios[nome_contato] = "ATENDIMENTO_HUMANO"
                                        tentativas_falhas.pop(nome_contato, None)
                                    else:
                                        resposta = "Desculpe, mas a imagem parece estar ilegível ou sem texto de erro. Por favor, tente enviar uma imagem mais clara ou informe o erro manualmente."
                                        tentativas_falhas[nome_contato] = falhas
                                
                                if os.path.exists(caminho_imagem):
                                    os.remove(caminho_imagem)

                            except Exception as e:
                                print(f"Ocorreu um erro ao processar a imagem: {e}")
                                resposta = "Desculpe, mas ocorreu um erro interno ao processar a imagem. Por favor, digite o erro manualmente."

                        elif mensagem_limpa:
                            print(f"O robô leu a mensagem: '{mensagem_limpa}'")
                            saudacoes = ["oi", "olá", "ola", "bom dia", "boa tarde", "boa noite", "preciso de ajuda", "menu", "deu erro"]

                            if mensagem_limpa in saudacoes:
                                    resposta = ("Olá! Sou o assistente virtual de suporte da Insoft.\n\n"
                                            "Para te ajudar de forma mais rápida, por favor escolha uma das opções abaixo:\n"
                                            "A - Nota fiscal (NFe)\n"
                                            "B - Vendas\n"
                                            "C - Outros Assuntos\n"
                                            "E - Encerrar atendimento")

                            elif mensagem_limpa in ["a", "nfe"]:
                                resposta = "Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre NFe."

                            elif mensagem_limpa in ["b", "vendas"]:
                                resposta = "Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre vendas."

                            elif mensagem_limpa in ["c", "outros"]:
                                resposta = "Por favor, informe o erro que aparece na tela ou a dúvida que você tem sobre outros assuntos."

                            elif mensagem_limpa in ["e", "sair"]:
                                resposta = "O seu atendimento está sendo encerrado.\n\nPara nos ajudar a melhorar, como você avalia o meu atendimento de *1 a 5*? (Sendo 1 Ruim e 5 Excelente)"
                                estado_usuarios[nome_contato] = "AGUARDANDO_AVALIACAO"
                            
                            else:
                                resposta = buscar_resposta(ultima_mensagem)
                                
                                if "Desculpe" in resposta:
                                    falhas = tentativas_falhas.get(nome_contato, 0) + 1
                                    if falhas >= 2:
                                        resposta = "Parece que não estou conseguindo encontrar a solução para o seu problema.\n\nEstou encaminhando o seu caso para um de nossos atendentes. Por favor, aguarde um momento."
                                        estado_usuarios[nome_contato] = "ATENDIMENTO_HUMANO"
                                        tentativas_falhas.pop(nome_contato, None)
                                    else:
                                        tentativas_falhas[nome_contato] = falhas
                                else:
                                    tentativas_falhas[nome_contato] = 0
                                    resposta += "\n\nEssa solução resolveu o seu problema? (Responda *Sim* ou *Não*)"
                                    estado_usuarios[nome_contato] = "AGUARDANDO_CONFIRMACAO"

                        else:
                            print("Formato não suportado")
                            falhas = tentativas_falhas.get(nome_contato, 0) + 1
                            if falhas >= 2:
                                resposta = "Como não estou conseguindo entender o formato enviado, estou encaminhando o seu caso para um de nossos atendentes. Por favor, aguarde um momento."
                                estado_usuarios[nome_contato] = "ATENDIMENTO_HUMANO"
                                tentativas_falhas.pop(nome_contato, None)
                            else:
                                resposta = "Desculpe, mas o tipo de mensagem que você enviou não é suportado pelo nosso robô. Por favor, envie uma mensagem de texto ou uma imagem clara do erro que você está enfrentando."
                                tentativas_falhas[nome_contato] = falhas

                    
                    caixa_texto_envio = navegador.find_element(By.XPATH, '//*[@id="main"]//footer//div[@contenteditable="true"]')

                    for linha in resposta.split('\n'):
                        linha_sem_emoji = "".join(c for c in linha if ord(c) <= 0xFFFF)
                        caixa_texto_envio.send_keys(linha_sem_emoji)
                        caixa_texto_envio.send_keys(Keys.SHIFT , Keys.ENTER)

                    caixa_texto_envio.send_keys(Keys.ENTER)

                    print(f"Resposta enviada com sucesso!")
                    
                    time.sleep(2)

                    acoes = ActionChains(navegador)
                    acoes.send_keys(Keys.ESCAPE).perform()

                else:
                    print("Abri a conversa, mas não consegui ler o texto da mensagem.")
                    acoes = ActionChains(navegador)
                    acoes.send_keys(Keys.ESCAPE).perform()

            except StaleElementReferenceException:
                print("A página do WhatsApp atualizou enquanto o robô lia. Recalculando...")
                acoes = ActionChains(navegador)
                acoes.send_keys(Keys.ESCAPE).perform()
                continue

    except TimeoutException:
        pass
    except Exception as e:
        print(f"Ocorreu um erro na automação: {e}")
        time.sleep(2)
          
    time.sleep(3)
