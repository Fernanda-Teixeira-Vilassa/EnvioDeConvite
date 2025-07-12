import pandas as pd
import pyautogui
import pyperclip
import time
import os

# Caminho do atalho do WhatsApp Desktop (ajustado)
caminho_whatsapp = r"C:\Users\nanda\OneDrive\Documents\Documentos\Envio de convite\WhatsApp.lnk"

# Mensagem base
mensagem_base = """*Queridos amigos e familiares,*

O nosso grande dia está chegando! 💍✨  
E será uma alegria imensa ter você com a gente celebrando esse momento tão especial.  

Criamos com muito carinho um site com todos os detalhes da cerimônia:  
🌐 *https://casamento-max-e-silva.vercel.app/*

No site você encontrará:
- Data, horário e localização da cerimônia 🗓📍  
- Informações sobre presentes 🎁  
- E um espaço para confirmar sua presença ✉  

*Como confirmar sua presença:*
1. Acesse o site  
2. Role até a seção “Confirmação de Presença”  
3. Informe a quantidade de pessoas e os nomes  
4. Clique em “Confirmar via WhatsApp”  

Se tiver qualquer dificuldade, é só falar com a gente diretamente!  
Sua presença é um presente precioso que vai tornar esse dia ainda mais inesquecível. 💚

Com muito carinho,  
*Max & Silva*"""

# Abre o WhatsApp Desktop
os.startfile(caminho_whatsapp)
time.sleep(12)  # Aguarda o app abrir

# Caminho da imagem que você capturou da caixa de envio
imagem_caixa_mensagem = r'caixa_mensagem.png'

# Função para localizar e clicar na caixa de mensagem
def clicar_caixa_mensagem():
    try:
        # Localiza a caixa de mensagem na tela
        caixa = pyautogui.locateOnScreen(imagem_caixa_mensagem, confidence=0.8)
        
        if caixa is not None:
            # Obtém o centro da caixa de mensagem
            centro_caixa = pyautogui.center(caixa)
            
            # Clica na caixa de mensagem
            pyautogui.click(centro_caixa)
            print("✅ Caixa de mensagem localizada e clicada!")
        else:
            print("❌ Não foi possível localizar a caixa de mensagem.")
    except Exception as e:
        print(f"❌ Erro ao tentar localizar a caixa de mensagem: {e}")

# Lê a planilha com os convidados
convidados = pd.read_excel("convidados.xlsx")

# Lista de falhas
falhas = []

# Função para buscar e enviar mensagem
def enviar_mensagem(numero, nome):
    try:
        # Clica na barra de pesquisa (ajuste se necessário para seu WhatsApp)
        pyautogui.click(x=200, y=150)
        time.sleep(1)

        # Digita o número com DDI
        numero_completo = f"{numero}"
        pyperclip.copy(numero_completo)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(2)
        pyautogui.press("enter")  # Entra na conversa
        time.sleep(3)

        # Chama a função para clicar na caixa de mensagem
        clicar_caixa_mensagem()
        time.sleep(1)

        # Cola e envia a mensagem
        pyperclip.copy(mensagem_base)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(1)
        pyautogui.press("enter")

        print(f"✅ Mensagem enviada para {nome} ({numero_completo})")

    except Exception as e:
        print(f"❌ Erro ao enviar para {nome} ({numero}): {e}")
        falhas.append({"Nome": nome, "Número": numero})

# Envia para todos os contatos
for index, row in convidados.iterrows():
    nome = row["Nome"]
    numero = str(row["Número"])
    enviar_mensagem(numero, nome)
    time.sleep(5)

# Registra as falhas
if falhas:
    pd.DataFrame(falhas).to_excel("falhas_envio.xlsx", index=False)
    print("\n⚠️ Erros registrados em 'falhas_envio.xlsx'")
else:
    print("\n🎉 Todas as mensagens foram enviadas com sucesso!")
