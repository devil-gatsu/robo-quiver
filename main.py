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
        caminho_imagem = os.path.join(os.getcwd(), nome_imagem)
        if not os.path.exists(caminho_imagem):
            print(f"[!] Aviso: Imagem '{nome_imagem}' não encontrada na pasta!")
            return False
        
        ponto = pyautogui.locateCenterOnScreen(caminho_imagem, confidence=confidence)
        if ponto:
            pyautogui.moveTo(ponto.x, ponto.y, duration=0.4)
            if cliques == 1:
                pyautogui.click()
            elif cliques == 2:
                pyautogui.doubleClick()
            time.sleep(esperar)
            return True
        else:
            return False
    except Exception as e:
        print(f"[Erro] Falha ao tentar clicar em {nome_imagem}: {e}")
        return False

def procurar_imagem_na_tela(nome_imagem, confidence=0.8):
    """Apenas procura a imagem para checagem, sem clicar."""
    caminho = os.path.join(os.getcwd(), nome_imagem)
    if os.path.exists(caminho):
        return pyautogui.locateOnScreen(caminho, confidence=confidence) is not None
    return False

def processar_planilha():
    arquivo_excel = 'Ocorrencias_MID.xlsx'
    
    if not os.path.exists(arquivo_excel):
        print(f"Erro crítico: O arquivo '{arquivo_excel}' não foi encontrado!")
        input("Pressione Enter para sair...")
        return

    wb = openpyxl.load_workbook(arquivo_excel)
    planilha = wb.active

    print("\n========================================================")
    print("ROBÔ INICIANDO EM 5 SEGUNDOS...")
    print("1. Feche o LibreOffice/Excel (se estiver aberto).")
    print("2. Puxe o Chrome para o MONITOR PRINCIPAL.")
    print("3. Clique no Chrome para deixá-lo focado.")
    print("========================================================\n")
    time.sleep(5)

    for linha in range(2, planilha.max_row + 1):
        apolice = planilha.cell(row=linha, column=1).value
        endosso = planilha.cell(row=linha, column=2).value
        ocorrencia = planilha.cell(row=linha, column=5).value
        
        # Pega a cor atual da célula (se já tiver)
        cor_atual = planilha.cell(row=linha, column=1).fill.start_color.index

        # 1. CONDIÇÃO DE PARADA: Se a coluna A estiver vazia
        if not apolice:
            break

        # 2. CONDIÇÃO DE PULO: Se já estiver pintado de Amarelo ou Azul
        if type(cor_atual) == str and ('FFFF00' in cor_atual or '00B0F0' in cor_atual):
            print(f"[{linha}] Apólice {apolice} já está pintada. Pulando...")
            continue

        # 3. CONDIÇÃO DE PULO: Se for Endosso (Coluna B preenchida)
        if endosso is not None and str(endosso).strip() != "":
            print(f"[{linha}] Apólice {apolice} é um Endosso (Col B={endosso}). Pulando...")
            continue

        ocorrencias_alvo = [
            "SOMA DOS PRÊMIOS DAS PARCS NÃO BATE COM O PRÊMIO TOTAL",
            "INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DA PROPOSTA",
            "INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DE EMISSÃO"
        ]

        if ocorrencia in ocorrencias_alvo:
            print(f"\n[{linha}] Iniciando Apólice: {apolice} | Ocorrência: {ocorrencia[:15]}...")
            
            pyperclip.copy(str(apolice))
            
            # Navegação no Quiver
            if not clicar_na_imagem('campo_pesquisa.png', cliques=2):
                print("Campo de pesquisa não encontrado. Marcando como Azul.")
                planilha.cell(row=linha, column=1).fill = AZUL
                continue
            
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.hotkey('ctrl', 'v')
            pyautogui.press('enter')
            time.sleep(3) 
            
            # Clica no ícone para abrir a apólice após pesquisar
            clicar_na_imagem('icone_abrir_apolice.png', esperar=3)

            # Clica na aba de Ocorrências
            if not clicar_na_imagem('botao_ocorrencias.png', esperar=2):
                print("Aba 'Ocorrências' não encontrada. Abortando esta linha.")
                planilha.cell(row=linha, column=1).fill = AZUL
                clicar_na_imagem('logomarca.png', esperar=3) # Volta pro início
                continue

            # 4. CONDIÇÃO DE PULO (Quiver): Verifica se já está Status 2 ou 3
            # A lógica é: se não achar a imagem do status 1, presume que já andou pra frente
            if not procurar_imagem_na_tela('status_1.png', confidence=0.85):
                print(f"Apólice {apolice} não possui Status 1. Provavelmente já é 2 ou 3. Pulando...")
                clicar_na_imagem('logomarca.png', esperar=3) # Volta pro início
                continue

            sucesso_cliques = False

            # === FLUXO A: SOMA DOS PRÊMIOS ===
            if ocorrencia == "SOMA DOS PRÊMIOS DAS PARCS NÃO BATE COM O PRÊMIO TOTAL":
                if clicar_na_imagem('aba_premios.png', esperar=2):
                    clicar_na_imagem('data_vencimento.png')
                    clicar_na_imagem('clicar_fora.png')
                    
                    # Tenta clicar no gravar. Se não achar, faz um scroll (rolar pra baixo) e tenta de novo
                    if não procurar_imagem_na_tela('gravar.png'):
                        pyautogui.scroll(-500) # Rola a página para baixo
                        time.sleep(1)
                    
                    if clicar_na_imagem('gravar.png', esperar=4):
                        sucesso_cliques = True
                else:
                    print("Aba 'Prêmios' não encontrada.")

            # === FLUXO B: VIGÊNCIA MAIOR/MENOR ===
            elif ocorrencia in ["INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DA PROPOSTA", 
                                "INÍCIO DE VIGÊNCIA ANTERIOR A 30 DIAS DA DATA DE EMISSÃO"]:
                # Como a vigência fica no topo, se o robô tiver descido a página antes, subimos de volta
                pyautogui.scroll(1000) 
                time.sleep(1)
                
                if clicar_na_imagem('data_proposta.png'):
                    clicar_na_imagem('clicar_fora.png')
                    
                    if não procurar_imagem_na_tela('gravar.png'):
                        pyautogui.scroll(-500)
                        time.sleep(1)
                        
                    if clicar_na_imagem('gravar.png', esperar=4):
                        sucesso_cliques = True

            # Validação Final de Status
            # Volta para aba ocorrências para checar
            pyautogui.scroll(1000) # Rola pra cima pra garantir que o botão está visível
            time.sleep(1)
            clicar_na_imagem('botao_ocorrencias.png', esperar=2)
            
            # Dá uma olhadinha pra baixo se necessário nas ocorrências
            pyautogui.scroll(-300)
            time.sleep(1)
            
            # Se ainda achar o "Status 1", é porque deu ruim
            ocorrencia_ainda_ativa = procurar_imagem_na_tela('status_1.png', confidence=0.85)

            # Colore a planilha
            if sucesso_cliques and not ocorrencia_ainda_ativa:
                print(f"Sucesso: Apólice {apolice} foi para status 3!")
                planilha.cell(row=linha, column=1).fill = AMARELO
            else:
                print(f"Pendente: Apólice {apolice} continuou no 1 ou erro ao gravar.")
                planilha.cell(row=linha, column=1).fill = AZUL
                
            wb.save('Ocorrencias_MID_Atualizado.xlsx')
            
            # Volta para o início (tela de pesquisa) para a próxima linha
            clicar_na_imagem('logomarca.png', esperar=3)

    print("\nPROCESSO FINALIZADO! Pode abrir o arquivo Ocorrencias_MID_Atualizado.xlsx.")
    input("Pressione Enter para fechar esta janela...")

if __name__ == "__main__":
    processar_planilha()
