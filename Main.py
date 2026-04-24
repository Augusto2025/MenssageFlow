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
# "system" faz o app seguir o tema (Claro/Escuro) do Windows automaticamente
ctk.set_appearance_mode("system") 
ctk.set_default_color_theme("blue")

excel_data = None
excel_nome = "Nenhum"
image_path = None
image_nome = "Nenhuma"
tempo_carregamento = 25 
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
    # Cor verde para sucesso, cinza para neutro (funciona em ambos os temas)
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
        relatorio += f"Nome: {erro['nome']} | Telefone: {erro['telefone']}\n"
    
    txt.insert("0.0", relatorio)
    txt.configure(state="disabled")

# --- NOVA FUNÇÃO DE LIMPEZA DE NÚMERO ---
def limpar_numero_br(telefone):
    # Remove tudo que não é dígito (parênteses, traços, espaços)
    n = ''.join(filter(str.isdigit, str(telefone)))
    
    # Se o número começar com 0, remove o zero (ex: 011...)
    if n.startswith('0'):
        n = n[1:]
        
    # Se o número não tem o DDI (55), nós adicionamos
    if not n.startswith('55'):
        # Se após tirar o 0 sobrarem 10 ou 11 dígitos, é um número BR sem 55
        if len(n) == 10 or len(n) == 11:
            n = '55' + n
            
    # Validação final: Um número BR com 55 + DDD + Numero deve ter 12 ou 13 dígitos
    if len(n) < 12 or len(n) > 13:
        return None # Número inválido ou internacional não suportado nesta regra
        
    return n

import unicodedata

# --- FUNÇÕES DE APOIO ---
def normalizar_texto(texto):
    """Remove acentos, deixa em minúsculo e limpa espaços."""
    if pd.isna(texto) or str(texto).strip() == "": return ""
    texto = str(texto).strip().lower()
    nfkd_form = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

import urllib.parse

# --- LÓGICA DE ENVIO "ZERO CLIQUES" ---
def processo_envio(msg):
    global lista_erros
    lista_erros = [] 
    sucessos = 0
    falhas = 0
    
    if 'nome' not in excel_data.columns or 'telefone' not in excel_data.columns:
        messagebox.showerror("Erro", "A tabela deve ter as colunas 'nome' e 'telefone'!")
        return

    total = len(excel_data)

    for i, row in excel_data.iterrows():
        nome_raw = row['nome']
        tel_raw = row['telefone']
        nome_normal = normalizar_texto(nome_raw)
        tel_str = str(tel_raw).strip()

        try:
            if nome_normal == "":
                raise Exception(f"Falta o NOME")
            if tel_str == "" or tel_str == "nan":
                raise Exception(f"Falta o NÚMERO")

            n_limpo = limpar_numero_br(tel_str)
            if n_limpo is None:
                raise Exception("Número fora do padrão BR")

            # 1. PREPARO DA MENSAGEM VIA URL (O WhatsApp já escreve para você)
            p_nome = nome_normal.split()[0].capitalize()
            mensagem_final = f"Olá {p_nome}! \n{msg}" if msg.strip() else ""
            # Codifica o texto para formato de URL (converte espaços e quebras de linha)
            msg_url = urllib.parse.quote(mensagem_final)

            # 2. ABRE O WHATSAPP JÁ COM O TEXTO ESCRITO
            link = f"https://web.whatsapp.com/send?phone={n_limpo}&text={msg_url}"
            webbrowser.open(link)
            
            # Espera o carregamento (importante ser generoso aqui)
            time.sleep(tempo_carregamento)

            # 3. ENVIO DA IMAGEM (CASO EXISTA)
            if image_path and os.path.exists(image_path):
                # Copia a imagem para o clipboard
                ps_command = f'Set-Clipboard -Path "{image_path}"'
                subprocess.run(['powershell', '-Command', ps_command], shell=True)
                time.sleep(2)
                
                # Cola a imagem (O WhatsApp vai colocar a imagem por cima do texto da URL)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(4) # Tempo para processar o anexo
                
                # Envia o conjunto (Imagem + Texto que já estava lá)
                pyautogui.press('enter')
            else:
                # Se for só texto, o texto já foi escrito pela URL, só dar Enter
                pyautogui.press('enter')

            # 4. SUCESSO
            sucessos += 1
            time.sleep(2) # Pequena pausa antes de fechar a aba
            pyautogui.hotkey('ctrl', 'w') # Fecha a aba para não poluir o navegador
            
        except Exception as e:
            falhas += 1
            lista_erros.append({"nome": str(nome_raw), "telefone": str(e)})
        
        progress_bar.set((i + 1) / total)
        time.sleep(tempo_intervalo)
        
    messagebox.showinfo("Finalizado", f"✅ Sucessos: {sucessos}\n❌ Falhas: {falhas}")


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

    # Removido fg_color fixo para que o frame se adapte ao fundo escuro/claro automaticamente
    main_frame = ctk.CTkScrollableFrame(root, height=630, width=450)
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

    # Removido fg_color e text_color fixos para o campo de texto se adaptar ao tema
    message_text = ctk.CTkTextbox(main_frame, height=100, border_width=1, border_color="#CED4DA")
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