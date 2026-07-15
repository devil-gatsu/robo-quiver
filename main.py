import openpyxl
from openpyxl.styles import PatternFill
import pyautogui
import pyperclip
import time
import os

# Trava de segurança: arraste o mouse para o canto da tela para abortar
pyautogui.FAILSAFE = True

AMARELO = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
AZUL = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')

def clicar_na_imagem(nome_imagem, cliques=1, esperar=1.5, confidence=0.8):
    try:
        caminho = os.path.join(os.getcwd(), nome_imagem)
        if not os.path.exists(caminho):
            return False
        
        ponto = pyautogui.locateCenterOnScreen(caminho, confidence=confidence)
        if ponto:
            pyautogui.moveTo(ponto.x, ponto.y, duration=0.3)
            time.sleep(0.2)
            if cliques == 1:
                pyautogui.click()
            elif cliques == 2:
                pyautogui.doubleClick()
            time.sleep(esperar)
            return True
        return False
    except:
        return False

def descer_e_clicar(nome_imagem, cliques=1, esperar=1.5, confidence=0.8, tentativas=15):
    """Usa a seta para baixo do teclado para rolar a tela suavemente até achar e clicar no botão"""
    for _ in range(tentativas):
        # Tenta clicar
        if clicar_na_imagem(nome_imagem, cliques, esperar, confidence):
            return True # Achou e clicou!
        
        # Se não achou, clica no canto da tela (pra focar na página) e desce um pouco com a setinha
        pyautogui.press('down', presses=4)
        time.sleep(0.5)
        
    print(f"[!] Não consegui encontrar o botão: {nome_imagem}")
    return False

def procurar_imagem_na_tela(nome_imagem, confidence=0.8):
    caminho = os.path.join(os.getcwd(), nome_imagem)
    if os.path.exists(caminho):
        return pyautogui.locateOnScreen(caminho, confidence=confidence) is not None
    return False

def processar_planilha():
    arquivo_excel = 'Ocorrencias_MID.xlsx'
    if not os.path.exists(arquivo_excel):
        print("Erro: Planilha Ocorrencias_MID.xlsx não encontrada!")
        input("Pressione Enter para sair...")
        return

    wb = openpyxl.load_workbook(arquivo_excel)
    planilha = wb.active

    print("\n========================================================")
    print("ROBÔ INICIANDO EM 5 SEGUNDOS...")
    print("Deixe o Chrome no Monitor Principal e na tela inicial.")
    print("========================================================\n")
    time.sleep(5)

    for linha in range(2, planilha.max_row + 1):
        apolice = planilha.cell(row=linha, column=1).value
        endosso = planilha.cell(row=linha, column=2).value
        ocorrencia = planilha.cell(row=linha, column=5).value
        
        cor_atual = planilha.cell(row=linha, column=1).fill.start_color.index

        if not apolice: break
        if type(cor_atual) == str and ('FFFF00' in cor_atual or '00B0F0' in cor_atual): continue
        if endosso is not None and str(endosso).strip() != "": continue

        ocorrencias_alvo = [
            "SOMA DOS PRÊMIOS DAS PARCS NÃO BATE COM O PRÊMIO TOTAL",
            "INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DA PROPOSTA",
            "INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DE EMISSÃO"
        ]

        if ocorrencia in ocorrencias_alvo:
            print(f"\n[{linha}] Apólice: {apolice}")
            
            # --- 1. PESQUISAR APÓLICE ---
            if not clicar_na_imagem('campo_pesquisa.png'):
                planilha.cell(row=linha, column=1).fill = AZUL
                continue
            
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyperclip.copy(str(apolice))
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('enter')
            time.sleep(4) 
            
            clicar_na_imagem('icone_abrir_apolice.png', esperar=3)
            
            # Clica num espaço neutro da página para garantir que as setinhas do teclado vão rolar a tela
            pyautogui.move(300, 0, duration=0.3)
            pyautogui.click() 
            time.sleep(1)

            # --- 2. BUSCAR E CLICAR NO BOTÃO OCORRÊNCIAS ---
            print(f"[{linha}] Descendo a tela procurando Ocorrências...")
            if not descer_e_clicar('botao_ocorrencias.png', esperar=2):
                planilha.cell(row=linha, column=1).fill = AZUL
                pyautogui.press('home') # Joga a tela pro topo
                clicar_na_imagem('logomarca.png', esperar=3)
                continue

            if not procurar_imagem_na_tela('status_1.png'):
                pyautogui.press('home')
                clicar_na_imagem('logomarca.png', esperar=3)
                continue

            sucesso_cliques = False

            if ocorrencia == "SOMA DOS PRÊMIOS DAS PARCS NÃO BATE COM O PRÊMIO TOTAL":
                if descer_e_clicar('aba_premios.png', esperar=2):
                    clicar_na_imagem('data_vencimento.png')
                    clicar_na_imagem('clicar_fora.png')
                    
                    if descer_e_clicar('gravar.png', esperar=4):
                        sucesso_cliques = True

            elif ocorrencia in ["INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DA PROPOSTA", 
                                "INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DE EMISSÃO"]:
                pyautogui.press('home') # Volta pro topo da página para achar a Proposta
                time.sleep(1)
                
                if clicar_na_imagem('data_proposta.png'):
                    clicar_na_imagem('clicar_fora.png')
                    
                    if descer_e_clicar('gravar.png', esperar=4):
                        sucesso_cliques = True

            # --- VERIFICAÇÃO FINAL ---
            pyautogui.press('home') # Volta pro topo para descer de novo até ocorrências
            time.sleep(1)
            descer_e_clicar('botao_ocorrencias.png', esperar=2)
            
            ocorrencia_ainda_ativa = procurar_imagem_na_tela('status_1.png')

            if sucesso_cliques and not ocorrencia_ainda_ativa:
                planilha.cell(row=linha, column=1).fill = AMARELO
            else:
                planilha.cell(row=linha, column=1).fill = AZUL
                
            wb.save('Ocorrencias_MID_Atualizado.xlsx')
            
            pyautogui.press('home') # Joga a tela pro topo para clicar na logomarca
            time.sleep(1)
            clicar_na_imagem('logomarca.png', esperar=3)

    print("\nPROCESSO FINALIZADO!")
    input("Pressione Enter para fechar...")

if __name__ == "__main__":
    processar_planilha()
