import pyautogui as pa
import openpyxl
from datetime import datetime
import time
import os
import re
import pytesseract
from PIL import Image
import pytesseract

# Caminho direto pro executável
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
print(pytesseract.get_tesseract_version())


# ========== CONFIGURAÇÕES INICIAIS ==========

caminho_planilha = r'H:\\Meu Drive\\Clientes\\UNION\\Bases\\BASE UNION.xlsx'
destino_certidoes = r'I:\Drives compartilhados\Operacional\1. Hefsys\2. CERTIDOES NEGATIVAS\RFB\2025-05'
destino_erros = r'C:\Users\Bruno Miguel\Desktop\Erros CNDFED'

# Configure o caminho do executável do Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pa.PAUSE = 1.0

# ========== FUNÇÃO PARA VERIFICAR MENSAGEM DE NÃO EMISSÃO ==========

def verificar_mensagem_negativa(razao_social):
    print("🔎 Verificando mensagem de resposta da RFB...")
    
    # Captura uma área da tela onde a mensagem costuma aparecer (ajuste se necessário)
    screenshot = pa.screenshot(region = (10, 150, 1200, 300))  # x, y, largura, altura
    texto = pytesseract.image_to_string(screenshot, lang='por')

    if "são insuficientes para a emissão" in texto:
        nome_erro = re.sub(r'[\\/*?:"<>|]', '_', f"{razao_social}-CNDFED-INDEFERIDA-{datetime.now().strftime('%m%Y')}")
        caminho_erro = os.path.join(destino_erros, nome_erro + '.png')
        screenshot.save(caminho_erro)
        print(f"⚠️ Certidão não emitida. Print salvo como: {nome_erro}.png")
        pa.click(x=72, y=393)  # botão 'Nova consulta'
        return True
    
    print("✅ Certidão disponível para emissão.")
    
    pa.PAUSE = 3.0

    return False

# ========== VERIFICAÇÕES INICIAIS ==========

if not os.path.isfile(caminho_planilha):
    print(f"Planilha não encontrada: {caminho_planilha}")
    exit(1)

if not os.path.exists(destino_erros):
    os.makedirs(destino_erros)

try:
    workbook = openpyxl.load_workbook(caminho_planilha)
    base_sheet = workbook['Base CND-Federal']
except Exception as e:
    print(f"Erro ao carregar a planilha: {e}")
    exit(1)

# ========== PROCESSAMENTO ==========

for linha in base_sheet.iter_rows(min_row=3, max_row=110):
    try:
        razao_social = str(linha[1].value).strip()
        cnpj = str(linha[3].value).strip()
        print(f"\n🔄 Processando: {razao_social} | CNPJ: {cnpj}")

        pa.click(x=77, y=392)
        pa.write(cnpj)
        time.sleep(1)

        pa.click(x=64, y=441)
        time.sleep(5)

        pa.click(x=79, y=301)
        time.sleep(15)

        if verificar_mensagem_negativa(razao_social):
            continue

        mes_ano = datetime.now().strftime("%m%Y")
        nome_arquivo = re.sub(r'[\\/*?:"<>|]', '_', f"{razao_social}-CNDFED-{mes_ano}")
        pa.write(nome_arquivo)
        time.sleep(1)

        pa.click(x=122, y=47)
        pa.write(destino_certidoes)
        pa.press('enter')
        time.sleep(1)

        pa.click(x=446, y=494)
        time.sleep(2)
        pa.click(x=102, y=554)

        print(f"✅ Certidão salva como: {nome_arquivo}.pdf")

        pa.click(x=75, y=373)

    except Exception as e:
        print(f"❌ Erro ao processar linha {linha[0].row}: {e}")
        continue

print("\n✅ Processo finalizado.")
