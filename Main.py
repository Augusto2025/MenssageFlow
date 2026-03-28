import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os, webbrowser, time, pyautogui, pyperclip, threading, json, subprocess, sys

# --- FUNÇÃO PARA RECURSOS NO EXE (ÍCONE) ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONFIGURAÇÕES GLOBAIS ---
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

excel_data = None
excel_nome = "Nenhum"
image_path = None
image_nome = "Nenhuma"
tempo_carregamento = 15 
tempo_intervalo = 30
arquivo_config = "config.json"
caminho_icone = resource_path(os.path.join("img", "MenssageFlow_icone.ico"))

# Variável para armazenar os detalhes das falhas
lista_erros = []

# --- FUNÇÕES DE SISTEMA ---
def obter_hwid():
    try:
        cmd = 'wmic csproduct get uuid'
        return str(subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip())
    except: return None

def validar_acesso():
    hwid_atual = obter_hwid()
    if not os.path.exists(arquivo_config):
        messagebox.showerror("Erro de Licença", "Arquivo 'config.json' não encontrado!")
        return False
    try:
        with open(arquivo_config, 'r') as f:
            dados = json.load(f)
        return dados.get("hwid_autorizado") == hwid_atual
    except: return False

def limpar_area_transferencia():
    try:
        subprocess.run('powershell Set-Clipboard -Value $null', shell=True)
        pyperclip.copy("")
    except: pass

# --- FUNÇÕES DA INTERFACE ---
def carregar_excel():
    global excel_data, excel_nome
    file = filedialog.askopenfilename(filetypes=[("Arquivos de Excel", "*.xlsx *.csv")])
    if file: 
        try:
            excel_data = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
            excel_nome = os.path.basename(file)
            atualizar_status()
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo: {e}")

def limpar_excel():
    global excel_data, excel_nome
    excel_data = None; excel_nome = "Nenhum"; atualizar_status()

def carregar_imagem():
    global image_path, image_nome
    file = filedialog.askopenfilename()
    if file: 
        image_path = os.path.abspath(file)
        image_nome = os.path.basename(file); atualizar_status()

def limpar_imagem():
    global image_path, image_nome
    image_path = None; image_nome = "Nenhuma"; atualizar_status()

def atualizar_status():
    status = f"Excel: {excel_nome}\nImagem: {image_nome}"
    info_label.configure(text=status, text_color="#28A745" if excel_data is not None else "#6C757D")

def mudar_tempo_carregar(v):
    global tempo_carregamento
    tempo_carregamento = int(v)
    lbl_carregar.configure(text=f"Espera de carregamento: {tempo_carregamento}s")

def mudar_tempo_intervalo(v):
    global tempo_intervalo
    tempo_intervalo = int(v)
    lbl_intervalo.configure(text=f"Intervalo entre contatos: {tempo_intervalo}s")

# Janela de detalhes dos erros
def mostrar_detalhes_erros():
    janela_erro = ctk.CTkToplevel()
    janela_erro.title("Relatório de Falhas")
    janela_erro.geometry("400x350")
    janela_erro.attributes("-topmost", True)
    
    txt = ctk.CTkTextbox(janela_erro, width=380, height=300)
    txt.pack(padx=10, pady=10)
    
    relatorio = "CONTATOS QUE FALHARAM:\n" + "="*30 + "\n"
    for erro in lista_erros:
        relatorio += f"Nome: {erro['nome']} | Número: {erro['numero']}\n"
    
    txt.insert("0.0", relatorio)
    txt.configure(state="disabled")

# --- LÓGICA DE ENVIO ---
def processo_envio(msg):
    global lista_erros
    lista_erros = [] # Limpa a lista para o novo envio
    sucessos = 0
    falhas = 0
    total = len(excel_data)
    largura, altura = pyautogui.size() 

    for i, row in excel_data.iterrows():
        # Captura os dados antes de tentar enviar para usar no log de erro se precisar
        nome_atual = str(row['nome']) if str(row['nome']) != "nan" else "Cliente"
        numero_atual = str(row['numero'])

        try:
            limpar_area_transferencia()
            n_limpo = ''.join(filter(str.isdigit, numero_atual))
            p_nome = nome_atual.strip().split()[0]
            mensagem_final = f"Olá {p_nome}! \n{msg}" if msg.strip() else ""

            webbrowser.open(f"https://web.whatsapp.com/send?phone={n_limpo}")
            time.sleep(tempo_carregamento)

            pyautogui.hotkey('alt', 'tab')
            time.sleep(1)
            pyautogui.click(largura/2, altura/2)
            time.sleep(1)

            if image_path and os.path.exists(image_path):
                ps_command = f'Set-Clipboard -Path "{image_path}"'
                subprocess.run(['powershell', '-Command', ps_command], shell=True)
                time.sleep(1.5)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(4) 
                if mensagem_final:
                    pyperclip.copy(mensagem_final)
                    time.sleep(0.5); pyautogui.hotkey('ctrl', 'v'); time.sleep(1)
                pyautogui.press('enter')
            elif mensagem_final:
                pyperclip.copy(mensagem_final)
                time.sleep(1); pyautogui.hotkey('ctrl', 'v'); time.sleep(1); pyautogui.press('enter')

            time.sleep(3); pyautogui.hotkey('ctrl', 'w')
            sucessos += 1
        except Exception:
            falhas += 1
            lista_erros.append({"nome": nome_atual, "numero": numero_atual})
        
        progress_bar.set((i + 1) / total)
        time.sleep(tempo_intervalo)
        
    # Mensagem final detalhada
    resultado_texto = f"Processo Finalizado!\n\n✅ Sucessos: {sucessos}\n❌ Falhas: {falhas}"
    messagebox.showinfo("Relatório MessageFlow", resultado_texto)
    
    if falhas > 0:
        if messagebox.askyesno("Ver Detalhes", "Deseja ver a lista de quem não recebeu a mensagem?"):
            mostrar_detalhes_erros()

def iniciar_envio():
    if excel_data is None: return messagebox.showerror("Erro", "Selecione o Excel!")
    msg = message_text.get("1.0", "end-1c")
    threading.Thread(target=processo_envio, args=(msg,), daemon=True).start()

# --- INTERFACE ---
if validar_acesso():
    root = ctk.CTk()
    root.title("MessageFlow")
    root.geometry("480x650")
    root.resizable(False, False)
    if os.path.exists(caminho_icone): root.iconbitmap(caminho_icone)

    main_frame = ctk.CTkScrollableFrame(root, fg_color="#F0F0F0", height=630, width=450)
    main_frame.pack(fill="both", expand=True, padx=15, pady=10)

    ctk.CTkLabel(main_frame, text="MessageFlow", font=("Segoe UI", 20, "bold"), text_color="#007BFF").pack(pady=(0, 15))
    ctk.CTkButton(main_frame, text="1. Iniciar Sessão", command=lambda: webbrowser.open("https://web.whatsapp.com"), height=32).pack(fill="x", pady=2)

    f_excel = ctk.CTkFrame(main_frame, fg_color="transparent")
    f_excel.pack(fill="x", pady=2)
    ctk.CTkButton(f_excel, text="2. Selecionar Excel", command=carregar_excel).pack(side="left", fill="x", expand=True)
    ctk.CTkButton(f_excel, text="Limpar", width=60, fg_color="#DC3545", command=limpar_excel).pack(side="right", padx=(5,0))

    f_img = ctk.CTkFrame(main_frame, fg_color="transparent")
    f_img.pack(fill="x", pady=2)
    ctk.CTkButton(f_img, text="3. Selecionar Imagem", command=carregar_imagem).pack(side="left", fill="x", expand=True)
    ctk.CTkButton(f_img, text="Limpar", width=60, fg_color="#DC3545", command=limpar_imagem).pack(side="right", padx=(5,0))

    info_label = ctk.CTkLabel(main_frame, text="Excel: Nenhum | Imagem: Nenhuma", font=("Segoe UI", 10), text_color="#6C757D", justify="left")
    info_label.pack(anchor="w", pady=5)

    message_text = ctk.CTkTextbox(main_frame, height=100, fg_color="#FFFFFF", border_width=1, border_color="#CED4DA", text_color="black")
    message_text.pack(fill="x", pady=10)

    lbl_carregar = ctk.CTkLabel(main_frame, text=f"Espera de carregamento: {tempo_carregamento}s", font=("Segoe UI", 10, "bold"), text_color="#007BFF")
    lbl_carregar.pack()
    slider_carregar = ctk.CTkSlider(main_frame, from_=5, to=60, command=mudar_tempo_carregar)
    slider_carregar.set(tempo_carregamento); slider_carregar.pack(fill="x", pady=5)

    lbl_intervalo = ctk.CTkLabel(main_frame, text=f"Intervalo entre contatos: {tempo_intervalo}s", font=("Segoe UI", 10, "bold"), text_color="#6F42C1")
    lbl_intervalo.pack()
    slider_intervalo = ctk.CTkSlider(main_frame, from_=5, to=120, command=mudar_tempo_intervalo)
    slider_intervalo.set(tempo_intervalo); slider_intervalo.pack(fill="x", pady=5)

    ctk.CTkButton(main_frame, text="▶️ Iniciar Envio", command=iniciar_envio, height=45, fg_color="#28A745").pack(fill="x", pady=15)
    progress_bar = ctk.CTkProgressBar(main_frame); progress_bar.pack(fill="x"); progress_bar.set(0)

    root.mainloop()