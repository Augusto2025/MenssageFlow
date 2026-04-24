import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os, webbrowser, time, pyautogui, threading, json, subprocess, sys, unicodedata, urllib.parse

# --- FUNÇÃO PARA RECURSOS NO EXE ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONFIGURAÇÕES GLOBAIS ---
ctk.set_appearance_mode("system") 
ctk.set_default_color_theme("blue")

TEMPO_CARREGAMENTO = 25 
TEMPO_POS_ENVIO = 6 
TEMPO_INTERVALO = 10
ARQUIVO_CONFIG = "config.json"
CAMINHO_ICONE = resource_path(os.path.join("img", "MenssageFlow_icone.ico"))

excel_data = None
excel_nome = "Nenhum"
image_path = None
image_nome = "Nenhuma"
lista_erros = []

# --- FUNÇÕES DE SISTEMA ---
def obter_hwid():
    try:
        cmd = 'wmic csproduct get uuid'
        return str(subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip())
    except: return None

def validar_acesso():
    hwid_atual = obter_hwid()
    if not os.path.exists(ARQUIVO_CONFIG):
        messagebox.showerror("Erro de Licença", "Arquivo 'config.json' não encontrado!")
        return False
    try:
        with open(ARQUIVO_CONFIG, 'r') as f:
            dados = json.load(f)
        return dados.get("hwid_autorizado") == hwid_atual
    except: return False

def limpar_numero_br(telefone):
    if pd.isna(telefone): return None
    n = ''.join(filter(str.isdigit, str(telefone)))
    if n.startswith('0'): n = n[1:]
    if not n.startswith('55'):
        if len(n) == 10 or len(n) == 11: n = '55' + n
    if len(n) < 12 or len(n) > 13: return None
    return n

def normalizar_texto(texto):
    if pd.isna(texto) or str(texto).strip() == "": return ""
    texto = str(texto).strip().lower()
    nfkd_form = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

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
            messagebox.showerror("Erro", f"Falha ao ler Excel: {e}")

def limpar_excel():
    global excel_data, excel_nome
    excel_data = None; excel_nome = "Nenhum"; atualizar_status()

def carregar_imagem():
    global image_path, image_nome
    file = filedialog.askopenfilename(filetypes=[("Imagens", "*.jpg;*.png;*.jpeg")])
    if file: 
        image_path = os.path.abspath(file)
        image_nome = os.path.basename(file); atualizar_status()

def limpar_imagem():
    global image_path, image_nome
    image_path = None; image_nome = "Nenhuma"; atualizar_status()

def atualizar_status():
    status = f"Excel: {excel_nome}\nImagem: {image_nome}"
    info_label.configure(text=status, text_color="#28A745" if excel_data is not None else "#6C757D")

def mostrar_detalhes_erros():
    janela_erro = ctk.CTkToplevel()
    janela_erro.title("Relatório")
    janela_erro.geometry("400x350")
    janela_erro.attributes("-topmost", True)
    txt = ctk.CTkTextbox(janela_erro, width=380, height=300)
    txt.pack(padx=10, pady=10)
    relatorio = "RELATÓRIO DE PROCESSAMENTO:\n" + "="*30 + "\n"
    for erro in lista_erros:
        relatorio += f"- {erro['nome']}: {erro['telefone']}\n"
    txt.insert("0.0", relatorio)
    txt.configure(state="disabled")

# --- LÓGICA DE ENVIO ---
def processo_envio(msg_base):
    global lista_erros
    lista_erros = [] 
    sucessos = 0
    falhas = 0
    
    df_temp = excel_data.copy()
    df_temp.columns = [normalizar_texto(c) for c in df_temp.columns]
    
    if 'nome' not in df_temp.columns or 'telefone' not in df_temp.columns:
        messagebox.showerror("Erro", "O Excel deve ter as colunas 'Nome' e 'Telefone'!")
        return

    total = len(df_temp)

    for i, row in df_temp.iterrows():
        nome_raw = row['nome']
        tel_raw = row['telefone']
        
        # Identificador para o relatório (usa o número se o nome falhar)
        id_relatorio = str(tel_raw) if not pd.isna(tel_raw) else f"Linha {i+1}"
        
        # Define o primeiro nome ou "Cliente"
        if pd.isna(nome_raw) or str(nome_raw).strip() == "":
            p_nome = "Cliente"
            lista_erros.append({"nome": "Aviso", "telefone": f"Número {id_relatorio} sem nome, enviado como 'Cliente'"})
        else:
            p_nome = str(nome_raw).split()[0].capitalize()

        try:
            n_limpo = limpar_numero_br(tel_raw)
            if not n_limpo:
                # Se tem nome mas não tem telefone
                nome_erro = p_nome if p_nome != "Cliente" else id_relatorio
                raise Exception(f"Faltou o telefone do Cliente {nome_erro}")

            # MONTAGEM DA MENSAGEM
            if "{nome}" in msg_base:
                mensagem_final = msg_base.replace("{nome}", p_nome)
            else:
                mensagem_final = f"Olá, {p_nome}!\n{msg_base}"

            msg_url = urllib.parse.quote(mensagem_final)

            link = f"https://web.whatsapp.com/send?phone={n_limpo}&text={msg_url}"
            webbrowser.open(link)
            
            time.sleep(TEMPO_CARREGAMENTO)

            if image_path and os.path.exists(image_path):
                ps_command = f'Set-Clipboard -Path "{image_path}"'
                subprocess.run(['powershell', '-Command', ps_command], shell=True)
                time.sleep(1.5)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(4) 
                pyautogui.press('enter')
            else:
                pyautogui.press('enter')

            time.sleep(TEMPO_POS_ENVIO) 
            sucessos += 1
            pyautogui.hotkey('ctrl', 'w') 
            
        except Exception as e:
            falhas += 1
            lista_erros.append({"nome": "ERRO", "telefone": str(e)})
        
        progress_bar.set((i + 1) / total)
        time.sleep(TEMPO_INTERVALO)
        
    messagebox.showinfo("Finalizado", f"✅ Sucessos: {sucessos}\n❌ Falhas/Avisos: {falhas}")
    if len(lista_erros) > 0: mostrar_detalhes_erros()

def iniciar_envio():
    if excel_data is None: return messagebox.showerror("Erro", "Selecione o Excel!")
    msg = message_text.get("1.0", "end-1c")
    if not msg.strip(): return messagebox.showwarning("Atenção", "Escreva uma mensagem!")
    threading.Thread(target=processo_envio, args=(msg,), daemon=True).start()

# --- INTERFACE ---
if validar_acesso():
    root = ctk.CTk()
    root.title("MessageFlow")
    root.geometry("480x580")
    root.resizable(False, False)
    if os.path.exists(CAMINHO_ICONE): root.iconbitmap(CAMINHO_ICONE)

    main_frame = ctk.CTkFrame(root, fg_color="transparent")
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(main_frame, text="MessageFlow", font=("Segoe UI", 22, "bold"), text_color="#007BFF").pack(pady=(0, 20))

    ctk.CTkButton(main_frame, text="1. Iniciar Sessão no WhatsApp Web", 
                  command=lambda: webbrowser.open("https://web.whatsapp.com"), 
                  height=35, fg_color="#333333", hover_color="#444444").pack(fill="x", pady=5)

    f_excel = ctk.CTkFrame(main_frame, fg_color="transparent")
    f_excel.pack(fill="x", pady=5)
    ctk.CTkButton(f_excel, text="2. Selecionar Excel", command=carregar_excel, height=35).pack(side="left", fill="x", expand=True)
    ctk.CTkButton(f_excel, text="❌", width=40, height=35, fg_color="#DC3545", command=limpar_excel).pack(side="right", padx=(5,0))

    f_img = ctk.CTkFrame(main_frame, fg_color="transparent")
    f_img.pack(fill="x", pady=5)
    ctk.CTkButton(f_img, text="3. Selecionar Imagem (Opcional)", command=carregar_imagem, height=35, fg_color="#6C757D").pack(side="left", fill="x", expand=True)
    ctk.CTkButton(f_img, text="❌", width=40, height=35, fg_color="#DC3545", command=limpar_imagem).pack(side="right", padx=(5,0))

    info_label = ctk.CTkLabel(main_frame, text="Excel: Nenhum | Imagem: Nenhuma", font=("Segoe UI", 11), text_color="#6C757D")
    info_label.pack(pady=10)

    ctk.CTkLabel(main_frame, text="Mensagem:", font=("Segoe UI", 12)).pack(anchor="w")
    message_text = ctk.CTkTextbox(main_frame, height=150, border_width=1, border_color="#CED4DA")
    message_text.pack(fill="x", pady=5)

    ctk.CTkButton(main_frame, text="🚀 INICIAR ENVIO EM MASSA", command=iniciar_envio, 
                  height=50, fg_color="#28A745", hover_color="#218838", font=("Segoe UI", 14, "bold")).pack(fill="x", pady=20)

    progress_bar = ctk.CTkProgressBar(main_frame)
    progress_bar.pack(fill="x")
    progress_bar.set(0)

    root.mainloop()