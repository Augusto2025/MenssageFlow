import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os, webbrowser, time, pyautogui, pyperclip, threading, json, subprocess

# Configurações de Aparência
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class AuthSystem:
    def __init__(self):
        self.arquivo_config = "config.json"

    def obter_hwid(self):
        try:
            cmd = 'wmic csproduct get uuid'
            uuid = str(subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip())
            return uuid
        except:
            return None

    def validar_acesso(self):
        hwid_atual = self.obter_hwid()
        if not os.path.exists(self.arquivo_config):
            messagebox.showerror("Erro de Licença", "Arquivo 'config.json' não encontrado!\nInsira sua licença na pasta.")
            return False
        try:
            with open(self.arquivo_config, 'r') as f:
                dados = json.load(f)
            if dados.get("hwid_autorizado") == hwid_atual:
                return True 
            else:
                messagebox.showerror("Acesso Negado", "Esta licença não pertence a este computador.")
                return False
        except:
            messagebox.showerror("Erro", "Falha ao ler o arquivo de licença.")
            return False

class MessageFlowApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("MessageFlow")
        self.root.geometry("480x520")
        self.root.resizable(False, False)
        
        self.excel_data = None
        self.excel_nome = "Nenhum"
        self.image_path = None
        self.image_nome = "Nenhuma"
        
        self.enviando = False
        self.tempo_espera = 30 
        self.setup_ui()
        
    def setup_ui(self):
        cor_fundo = "#F0F0F0"
        self.root.configure(fg_color=cor_fundo)
        self.main_frame = ctk.CTkScrollableFrame(self.root, fg_color=cor_fundo, height=580, width=450)
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        ctk.CTkLabel(self.main_frame, text="MessageFlow", font=("Segoe UI", 20, "bold"), text_color="#007BFF").pack(pady=(0, 15))
        
        # 1. Botão Sessão
        ctk.CTkButton(self.main_frame, text="1. Iniciar Sessão", command=lambda: webbrowser.open("https://web.whatsapp.com"), height=32).pack(fill="x", pady=2)
        
        # 2. Frame Excel
        f_excel = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        f_excel.pack(fill="x", pady=2)
        ctk.CTkButton(f_excel, text="2. Selecionar Excel", command=self.carregar_excel).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_excel, text="Limpar", width=60, fg_color="#DC3545", command=self.limpar_excel).pack(side="right", padx=(5,0))
        
        # 3. Frame Imagem
        f_img = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        f_img.pack(fill="x", pady=2)
        ctk.CTkButton(f_img, text="3. Selecionar Imagem", command=self.carregar_imagem).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(f_img, text="Limpar", width=60, fg_color="#DC3545", command=self.limpar_imagem).pack(side="right", padx=(5,0))
        
        # Label de Status (Info Label Atualizada)
        self.info_label = ctk.CTkLabel(self.main_frame, text="Excel: Nenhum | Imagem: Nenhuma", font=("Segoe UI", 10), text_color="#6C757D", justify="left")
        self.info_label.pack(anchor="w", pady=5)

        self.message_text = ctk.CTkTextbox(self.main_frame, height=120, fg_color="#FFFFFF", border_width=1, border_color="#CED4DA", text_color="black")
        self.message_text.pack(fill="x", pady=10)
        
        self.tempo_label = ctk.CTkLabel(self.main_frame, text=f"Tempo entre mensagens: {self.tempo_espera}s", font=("Segoe UI", 10, "bold"), text_color="#6F42C1")
        self.tempo_label.pack()
        self.tempo_slider = ctk.CTkSlider(self.main_frame, from_=10, to=120, command=self.mudar_tempo)
        self.tempo_slider.set(30)
        self.tempo_slider.pack(fill="x", pady=5)

        self.send_button = ctk.CTkButton(self.main_frame, text="▶️ Iniciar Envio", command=self.iniciar_envio, height=45, fg_color="#28A745")
        self.send_button.pack(fill="x", pady=15)
        
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

    def atualizar_status(self):
        status = f"Excel: {self.excel_nome}\nImagem: {self.image_nome}"
        self.info_label.configure(text=status)
        # Muda a cor se o excel estiver carregado
        if self.excel_data is not None:
            self.info_label.configure(text_color="#28A745")
        else:
            self.info_label.configure(text_color="#6C757D")

    def carregar_excel(self):
        file = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.csv")])
        if file:
            self.excel_data = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
            self.excel_nome = os.path.basename(file)
            self.atualizar_status()

    def limpar_excel(self):
        self.excel_data = None
        self.excel_nome = "Nenhum (Removido)"
        self.atualizar_status()

    def carregar_imagem(self):
        file = filedialog.askopenfilename(filetypes=[("Imagens", "*.png *.jpg *.jpeg")])
        if file:
            self.image_path = file
            self.image_nome = os.path.basename(file)
            self.atualizar_status()

    def limpar_imagem(self):
        self.image_path = None
        self.image_nome = "Nenhuma (Removida)"
        self.atualizar_status()

    def mudar_tempo(self, v):
        self.tempo_espera = int(v)
        self.tempo_label.configure(text=f"Tempo entre mensagens: {self.tempo_espera}s")

    def iniciar_envio(self):
        if self.excel_data is None:
            return messagebox.showerror("Erro", "Selecione o Excel!")
        threading.Thread(target=self.processo_envio, args=(self.message_text.get("1.0", "end-1c"),), daemon=True).start()

    def processo_envio(self, msg):
        total = len(self.excel_data)
        for i, row in self.excel_data.iterrows():
            p_nome = str(row['nome']).strip().split()[0] if str(row['nome']) else "Cliente"
            n_limpo = ''.join(filter(str.isdigit, str(row['numero'])))

            pyperclip.copy(f"Olá {p_nome}! {msg}")
            webbrowser.open(f"https://web.whatsapp.com/send?phone={n_limpo}")
            
            time.sleep(15)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(2)
            pyautogui.hotkey('ctrl', 'w')
            
            self.progress_bar.set((i + 1) / total)
            time.sleep(self.tempo_espera)
        messagebox.showinfo("Fim", "Processo concluído!")

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    auth = AuthSystem()
    if auth.validar_acesso():
        app = MessageFlowApp()
        app.run()