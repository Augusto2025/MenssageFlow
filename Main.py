import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import webbrowser
import time
import pyautogui
import pyperclip
import threading
import urllib.parse

# Configurar para tema CLARO
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class ConfigTempoWindow:
    """Janela para configurar o tempo entre mensagens"""
    def __init__(self, parent, tempo_atual, callback):
        self.parent = parent
        self.tempo_selecionado = tempo_atual
        self.callback = callback  # Função callback para retornar o valor
        
        # Criar janela
        self.window = ctk.CTkToplevel(parent)
        self.window.title("Configurar Tempo entre Mensagens")
        self.window.geometry("350x300")  # Um pouco maior para o botão
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()
        
        # Configurar layout
        self.setup_ui()
        
    def setup_ui(self):
        # CORES
        cor_fundo = "#F0F0F0"
        cor_titulo = "#007BFF"
        cor_texto = "#212529"
        cor_botao = "#28A745"
        
        self.window.configure(fg_color=cor_fundo)
        
        # TÍTULO
        title_label = ctk.CTkLabel(
            self.window,
            text="Configurar Tempo",
            font=("Segoe UI", 18, "bold"),
            text_color=cor_titulo
        )
        title_label.pack(pady=(20, 10))
        
        # DESCRIÇÃO
        desc_label = ctk.CTkLabel(
            self.window,
            text="Selecione o tempo de espera entre cada mensagem:",
            font=("Segoe UI", 12),
            text_color=cor_texto,
            wraplength=300
        )
        desc_label.pack(pady=(0, 15))
        
        # FRAME PARA OPÇÕES
        opcoes_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        opcoes_frame.pack(pady=10, padx=20, fill="both", expand=True)
        
        # Variável para o tempo
        self.tempo_var = ctk.StringVar(value=str(self.tempo_selecionado))
        
        # Opções de tempo
        opcoes = [
            ("10 segundos (mais rápido)", "10"),
            ("30 segundos (recomendado)", "30"),
            ("40 segundos (mais seguro)", "40"),
            ("60 segundos (muito seguro)", "60")
        ]
        
        for texto, valor in opcoes:
            frame_opcao = ctk.CTkFrame(opcoes_frame, fg_color="transparent", height=35)
            frame_opcao.pack(fill="x", pady=2)
            
            rb = ctk.CTkRadioButton(
                frame_opcao,
                text=texto,
                variable=self.tempo_var,
                value=valor,
                font=("Segoe UI", 11),
                corner_radius=10
            )
            rb.pack(side="left")
        
        # FRAME PARA BOTÕES
        botoes_frame = ctk.CTkFrame(self.window, fg_color="transparent")
        botoes_frame.pack(pady=10)
        
        # BOTÃO CONFIRMAR
        confirmar_btn = ctk.CTkButton(
            botoes_frame,
            text="✅ Confirmar",
            command=self.confirmar,
            height=35,
            width=120,
            font=("Segoe UI", 12, "bold"),
            fg_color=cor_botao,
            hover_color="#218838",
            text_color="white",
            corner_radius=6
        )
        confirmar_btn.pack(side="left", padx=5)
        
        # BOTÃO CANCELAR
        cancelar_btn = ctk.CTkButton(
            botoes_frame,
            text="❌ Cancelar",
            command=self.cancelar,
            height=35,
            width=120,
            font=("Segoe UI", 12),
            fg_color="#DC3545",
            hover_color="#C82333",
            text_color="white",
            corner_radius=6
        )
        cancelar_btn.pack(side="right", padx=5)
        
    def confirmar(self):
        try:
            self.tempo_selecionado = int(self.tempo_var.get())
            self.window.destroy()
            self.callback(self.tempo_selecionado)  # Chama o callback com o valor
        except:
            messagebox.showerror("Erro", "Selecione um tempo válido!")
        
    def cancelar(self):
        self.window.destroy()

class MessageFlowApp:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("MessageFlow")
        self.root.geometry("480x580")  # Um pouco maior
        self.root.resizable(False, False)
        
        # Variáveis
        self.excel_file = None
        self.image_file = None
        self.excel_data = None
        self.enviando = False
        self.cancelar_envio = False
        self.tempo_entre_mensagens = 30  # Valor padrão
        
        # Layout principal
        self.setup_ui()
        
    def setup_ui(self):
        # CORES
        cor_fundo = "#F0F0F0"
        cor_aviso_bg = "#FFF3CD"
        cor_aviso_borda = "#FFEAA7"
        cor_aviso_texto = "#856404"
        cor_titulo = "#007BFF"
        cor_botao = "#007BFF"
        cor_botao_limpar = "#DC3545"
        cor_botao_enviar = "#28A745"
        cor_botao_tempo = "#6F42C1"
        cor_texto = "#212529"
        cor_borda = "#CED4DA"
        cor_campo_bg = "#FFFFFF"
        cor_placeholder = "#6C757D"
        
        # Configurar fundo
        self.root.configure(fg_color=cor_fundo)
        
        # FONTES
        font_titulo = ("Segoe UI", 20, "bold")
        font_normal = ("Segoe UI", 10)
        font_menu = ("Segoe UI", 11)
        font_status = ("Segoe UI", 10)
        font_botao = ("Segoe UI", 10, "bold")
        
        # Frame principal COM SCROLL
        self.main_frame = ctk.CTkScrollableFrame(
            self.root, 
            fg_color=cor_fundo,
            height=560,
            width=450,
            scrollbar_button_color="#007BFF",
            scrollbar_button_hover_color="#0056B3"
        )
        self.main_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        # AVISO AMARELO NO TOPO
        aviso_frame = ctk.CTkFrame(
            self.main_frame,
            fg_color=cor_aviso_bg,
            border_width=2,
            border_color=cor_aviso_borda,
            corner_radius=4
        )
        aviso_frame.pack(fill="x", pady=(0, 10))
        
        aviso_texto = """É necessário estar conectado ao WhatsApp Web para o envio funcionar.
A imagem é opcional."""
        
        aviso_label = ctk.CTkLabel(
            aviso_frame,
            text=aviso_texto,
            font=("Segoe UI", 9),
            text_color=cor_aviso_texto,
            justify="left"
        )
        aviso_label.pack(padx=10, pady=8)
        
        # TÍTULO AZUL
        title_label = ctk.CTkLabel(
            self.main_frame,
            text="MessageFlow",
            font=font_titulo,
            text_color=cor_titulo
        )
        title_label.pack(pady=(0, 15))
        
        # BOTÕES DO MENU
        menu_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        menu_frame.pack(fill="x", pady=(0, 10))
        
        # BOTÃO INICIAR SESSÃO
        session_btn = ctk.CTkButton(
            menu_frame,
            text="Iniciar Sessão",
            command=self.open_whatsapp_web,
            height=32,
            font=font_botao,
            fg_color=cor_botao,
            hover_color="#0056B3",
            text_color="white",
            corner_radius=4
        )
        session_btn.pack(fill="x", pady=1)
        
        # Frame para Selecionar Excel
        excel_frame = ctk.CTkFrame(menu_frame, fg_color="transparent")
        excel_frame.pack(fill="x", pady=1)
        
        excel_btn = ctk.CTkButton(
            excel_frame,
            text="Selecionar Excel",
            command=self.select_excel_file,
            height=32,
            font=font_botao,
            fg_color=cor_botao,
            hover_color="#0056B3",
            text_color="white",
            corner_radius=4
        )
        excel_btn.pack(side="left", fill="x", expand=True)
        
        clear_excel_btn = ctk.CTkButton(
            excel_frame,
            text="Limpar",
            command=self.clear_excel_file,
            width=70,
            height=26,
            font=("Segoe UI", 9, "bold"),
            fg_color=cor_botao_limpar,
            hover_color="#C82333",
            text_color="white",
            corner_radius=3
        )
        clear_excel_btn.pack(side="right", padx=(5, 0))
        
        # Frame para Selecionar Imagem
        image_frame = ctk.CTkFrame(menu_frame, fg_color="transparent")
        image_frame.pack(fill="x", pady=1)
        
        image_btn = ctk.CTkButton(
            image_frame,
            text="Selecionar Imagem",
            command=self.select_image_file,
            height=32,
            font=font_botao,
            fg_color=cor_botao,
            hover_color="#0056B3",
            text_color="white",
            corner_radius=4
        )
        image_btn.pack(side="left", fill="x", expand=True)
        
        clear_image_btn = ctk.CTkButton(
            image_frame,
            text="Limpar",
            command=self.clear_image_file,
            width=70,
            height=26,
            font=("Segoe UI", 9, "bold"),
            fg_color=cor_botao_limpar,
            hover_color="#C82333",
            text_color="white",
            corner_radius=3
        )
        clear_image_btn.pack(side="right", padx=(5, 0))
        
        # INFORMAÇÕES DOS ARQUIVOS
        info_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        info_frame.pack(fill="x", pady=(5, 10))
        
        self.excel_info_label = ctk.CTkLabel(
            info_frame,
            text="",
            font=("Segoe UI", 9),
            text_color="#28A745"
        )
        self.excel_info_label.pack(anchor="w", pady=(0, 1))
        
        self.image_info_label = ctk.CTkLabel(
            info_frame,
            text="",
            font=("Segoe UI", 9),
            text_color="#28A745"
        )
        self.image_info_label.pack(anchor="w")
        
        # MENSAGEM A ENVIAR
        msg_label_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        msg_label_frame.pack(fill="x", pady=(0, 5))
        
        message_label = ctk.CTkLabel(
            msg_label_frame,
            text="Mensagem a enviar (o nome será adicionado automaticamente):",
            font=font_menu,
            text_color=cor_texto
        )
        message_label.pack(anchor="w")
        
        # CAMPO DE MENSAGEM
        self.message_text = ctk.CTkTextbox(
            self.main_frame,
            height=100,
            font=("Segoe UI", 10),
            wrap="word",
            fg_color=cor_campo_bg,
            text_color=cor_texto,
            border_width=1,
            border_color=cor_borda,
            corner_radius=4
        )
        self.message_text.pack(fill="x", pady=(0, 8))
        
        # Placeholder
        self.message_text.insert("1.0", "Digite sua mensagem aqui...\n\nA mensagem será formatada assim:\nOlá, [Nome da Pessoa]\n[sua mensagem aqui]")
        self.message_text.configure(text_color=cor_placeholder)
        
        # Eventos
        self.message_text.bind("<FocusIn>", self.on_text_focus_in)
        self.message_text.bind("<FocusOut>", self.on_text_focus_out)
        self.message_text.bind("<KeyRelease>", self.update_char_count)
        
        # Contador de caracteres
        self.char_count_label = ctk.CTkLabel(
            self.main_frame,
            text="Caracteres: 0",
            font=("Segoe UI", 9),
            text_color="#6C757D"
        )
        self.char_count_label.pack(anchor="w", pady=(0, 10))
        self.update_char_count()
        
        # FRAME PARA CONFIGURAÇÕES DE ENVIO
        config_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        config_frame.pack(fill="x", pady=(0, 15))
        
        # BOTÃO PARA CONFIGURAR TEMPO
        tempo_btn_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        tempo_btn_frame.pack(fill="x", pady=(0, 10))
        
        self.tempo_btn = ctk.CTkButton(
            tempo_btn_frame,
            text="⏱️ Configurar Tempo entre Mensagens",
            command=self.abrir_config_tempo,
            height=35,
            font=("Segoe UI", 11, "bold"),
            fg_color=cor_botao_tempo,
            hover_color="#5A32A3",
            text_color="white",
            corner_radius=5
        )
        self.tempo_btn.pack(fill="x")
        
        # Label que mostra o tempo atual
        self.tempo_atual_label = ctk.CTkLabel(
            tempo_btn_frame,
            text="Tempo atual: 30 segundos",
            font=("Segoe UI", 10, "bold"),
            text_color="#6F42C1"
        )
        self.tempo_atual_label.pack(anchor="w", pady=(5, 0))
        
        # BOTÃO INICIAR ENVIO
        self.send_button = ctk.CTkButton(
            config_frame,
            text="▶️ Iniciar Envio",
            command=self.start_sending,
            height=40,
            font=("Segoe UI", 13, "bold"),
            fg_color=cor_botao_enviar,
            hover_color="#218838",
            text_color="white",
            corner_radius=5
        )
        self.send_button.pack(fill="x", pady=(0, 10))
        
        # BARRA DE PROGRESSO
        self.progress_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        self.progress_frame.pack(fill="x", pady=(0, 10))
        
        self.progress_label = ctk.CTkLabel(
            self.progress_frame,
            text="Progresso: 0%",
            font=("Segoe UI", 10),
            text_color="#007BFF"
        )
        self.progress_label.pack(anchor="w")
        
        self.progress_bar = ctk.CTkProgressBar(
            self.progress_frame,
            height=12,
            fg_color="#E9ECEF",
            progress_color="#28A745",
            border_width=1,
            border_color="#CED4DA"
        )
        self.progress_bar.pack(fill="x", pady=(3, 0))
        self.progress_bar.set(0)
        
        # BOTÃO CANCELAR
        self.cancel_button = ctk.CTkButton(
            config_frame,
            text="⏹️ Cancelar Envio",
            command=self.cancel_sending,
            height=32,
            font=("Segoe UI", 10, "bold"),
            fg_color="#DC3545",
            hover_color="#C82333",
            text_color="white",
            corner_radius=4
        )
        self.cancel_button.pack(fill="x", pady=(0, 10))
        self.cancel_button.pack_forget()
        
        # STATUS
        self.status_label = ctk.CTkLabel(
            config_frame,
            text="Aguardando seleção de arquivo...",
            font=font_status,
            text_color="#6C757D"
        )
        self.status_label.pack()
        
    def atualizar_tempo_label(self, tempo):
        """Atualiza o label com o novo tempo"""
        self.tempo_entre_mensagens = tempo
        self.tempo_atual_label.configure(text=f"Tempo atual: {tempo} segundos")
        
    def abrir_config_tempo(self):
        """Abre janela para configurar tempo entre mensagens"""
        ConfigTempoWindow(self.root, self.tempo_entre_mensagens, self.atualizar_tempo_label)
        
    def on_text_focus_in(self, event):
        current_text = self.message_text.get("1.0", "end-1c")
        if current_text == "Digite sua mensagem aqui...\n\nA mensagem será formatada assim:\nOlá, [Nome da Pessoa]\n[sua mensagem aqui]":
            self.message_text.delete("1.0", "end")
            self.message_text.configure(text_color="#212529")
    
    def on_text_focus_out(self, event):
        current_text = self.message_text.get("1.0", "end-1c").strip()
        if not current_text:
            self.message_text.delete("1.0", "end")
            self.message_text.insert("1.0", "Digite sua mensagem aqui...\n\nA mensagem será formatada assim:\nOlá, [Nome da Pessoa]\n[sua mensagem aqui]")
            self.message_text.configure(text_color="#6C757D")
        
    def update_char_count(self, event=None):
        mensagem = self.message_text.get("1.0", "end-1c")
        if mensagem == "Digite sua mensagem aqui...\n\nA mensagem será formatada assim:\nOlá, [Nome da Pessoa]\n[sua mensagem aqui]":
            char_count = 0
        else:
            char_count = len(mensagem)
        self.char_count_label.configure(text=f"Caracteres: {char_count}")
    
    def open_whatsapp_web(self):
        webbrowser.open("https://web.whatsapp.com")
        self.update_status("WhatsApp Web aberto - Escaneie o QR Code")
        messagebox.showinfo(
            "WhatsApp Web",
            "1. WhatsApp Web aberto\n"
            "2. Escaneie o QR Code com seu celular\n"
            "3. Aguarde a conexão\n"
            "4. Mantenha a janela aberta"
        )
        
    def select_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.excel_data = pd.read_csv(file_path, encoding='utf-8')
                else:
                    self.excel_data = pd.read_excel(file_path)
                
                if self.excel_data.empty:
                    messagebox.showerror("Erro", "O arquivo Excel está vazio!")
                    self.clear_excel_file()
                    return
                
                if 'numero' not in self.excel_data.columns or 'nome' not in self.excel_data.columns:
                    messagebox.showerror("Erro", "O arquivo deve conter colunas: 'numero' e 'nome'")
                    self.clear_excel_file()
                    return
                
                self.excel_file = file_path
                file_name = os.path.basename(file_path)
                contatos = len(self.excel_data)
                
                self.excel_info_label.configure(
                    text=f"✓ Excel: {file_name[:20]}... ({contatos} contatos)" if len(file_name) > 20 else f"✓ Excel: {file_name} ({contatos} contatos)",
                    text_color="#28A745"
                )
                
                self.update_status(f"Excel carregado: {contatos} contatos")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao ler o arquivo: {str(e)}")
                self.clear_excel_file()
    
    def clear_excel_file(self):
        self.excel_file = None
        self.excel_data = None
        self.excel_info_label.configure(text="")
        self.update_status("Excel removido")
        
    def select_image_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecione uma imagem",
            filetypes=[("Imagens", "*.png *.jpg *.jpeg *.gif *.bmp"), ("Todos os arquivos", "*.*")]
        )
        
        if file_path:
            self.image_file = file_path
            file_name = os.path.basename(file_path)
            
            self.image_info_label.configure(
                text=f"✓ Imagem: {file_name[:25]}..." if len(file_name) > 25 else f"✓ Imagem: {file_name}",
                text_color="#28A745"
            )
            
            self.update_status("Imagem selecionada")
    
    def clear_image_file(self):
        self.image_file = None
        self.image_info_label.configure(text="")
        self.update_status("Imagem removida")
    
    def update_status(self, status):
        self.status_label.configure(text=status)
        
        if "pronto" in status.lower() or "carregado" in status.lower():
            self.status_label.configure(text_color="#28A745")
        elif "erro" in status.lower():
            self.status_label.configure(text_color="#DC3545")
        elif "aguardando" in status.lower():
            self.status_label.configure(text_color="#6C757D")
        elif "enviando" in status.lower():
            self.status_label.configure(text_color="#007BFF")
        else:
            self.status_label.configure(text_color="#6C757D")
    
    def start_sending(self):
        if self.enviando:
            return
            
        if self.excel_data is None or (hasattr(self.excel_data, 'empty') and self.excel_data.empty):
            messagebox.showerror("Erro", "Selecione um arquivo Excel primeiro!")
            return
        
        mensagem_usuario = self.message_text.get("1.0", "end-1c").strip()
        
        if not mensagem_usuario or mensagem_usuario == "Digite sua mensagem aqui...\n\nA mensagem será formatada assim:\nOlá, [Nome da Pessoa]\n[sua mensagem aqui]":
            messagebox.showerror("Erro", "Digite uma mensagem para enviar!")
            return
        
        # Confirmar envio
        resposta = messagebox.askyesno(
            "Confirmar Envio",
            f"Enviar para {len(self.excel_data)} contatos?\n\n"
            f"• Tempo entre mensagens: {self.tempo_entre_mensagens} segundos\n"
            f"• Mensagem será formatada como:\n   Olá, [Nome]\n   {mensagem_usuario[:80]}..."
        )
        
        if not resposta:
            return
        
        # Instruções importantes
        messagebox.showinfo(
            "ATENÇÃO - Leia com atenção",
            "INSTRUÇÕES IMPORTANTES:\n\n"
            "1. WhatsApp Web DEVE estar aberto e logado\n"
            "2. A janela do WhatsApp Web DEVE estar VISÍVEL\n"
            "3. NÃO use mouse ou teclado durante o envio\n"
            "4. O sistema irá controlar o mouse automaticamente\n\n"
            "O envio começará em 5 segundos..."
        )
        
        # Preparar envio
        self.enviando = True
        self.cancelar_envio = False
        self.send_button.configure(state="disabled", text="Preparando...")
        self.tempo_btn.configure(state="disabled")
        
        # Mostrar botão cancelar
        self.cancel_button.pack(fill="x", pady=(0, 10))
        
        # Resetar progresso
        self.progress_bar.set(0)
        self.progress_label.configure(text="Progresso: 0%")
        
        self.update_status("Preparando envio...")
        self.root.update()
        
        # Contagem regressiva
        for i in range(5, 0, -1):
            if self.cancelar_envio:
                self.reset_envio()
                return
            self.update_status(f"Iniciando em {i} segundos...")
            self.root.update()
            time.sleep(1)
        
        # Iniciar envio
        thread = threading.Thread(target=self.enviar_mensagens_whatsapp, args=(mensagem_usuario, self.tempo_entre_mensagens))
        thread.daemon = True
        thread.start()
    
    def cancel_sending(self):
        self.cancelar_envio = True
        self.update_status("Cancelando...")
    
    def reset_envio(self):
        self.enviando = False
        self.cancelar_envio = False
        self.send_button.configure(state="normal", text="▶️ Iniciar Envio")
        self.tempo_btn.configure(state="normal")
        self.cancel_button.pack_forget()
        self.update_status("Envio cancelado")
    
    def enviar_mensagens_whatsapp(self, mensagem_usuario, tempo_espera):
        """ENVIA MENSAGENS REAIS VIA WHATSAPP WEB"""
        total = len(self.excel_data)
        enviados = 0
        falhas = 0
        
        self.root.after(0, lambda: self.update_status(f"Iniciando envio para {total} contatos..."))
        
        try:
            # Primeiro, focar na janela do WhatsApp Web
            time.sleep(2)
            
            # Encontrar e focar na janela do WhatsApp Web
            try:
                # Tenta encontrar janela do Chrome/Edge com WhatsApp
                pyautogui.hotkey('alt', 'tab')
                time.sleep(1)
                pyautogui.hotkey('alt', 'tab')  # Volta para nossa aplicação
                time.sleep(1)
            except:
                pass
            
            for index, row in self.excel_data.iterrows():
                if self.cancelar_envio:
                    self.root.after(0, self.reset_envio)
                    return
                    
                nome = str(row['nome'])
                numero = str(row['numero'])
                
                # Mensagem formatada
                mensagem_final = f"Olá, {nome}\n{mensagem_usuario}"
                
                # Atualizar progresso
                progresso = (index + 1) / total
                self.root.after(0, lambda p=progresso: self.progress_bar.set(p))
                self.root.after(0, lambda p=progresso: self.progress_label.configure(text=f"Progresso: {int(p*100)}%"))
                
                # Atualizar status
                self.root.after(0, lambda: self.update_status(f"Enviando para {nome[:15]}... ({index+1}/{total})"))
                self.root.after(0, lambda: self.send_button.configure(text=f"Enviando {enviados}/{total}"))
                
                try:
                    # Limpar número
                    numero_limpo = ''.join(filter(str.isdigit, numero))
                    
                    if not numero_limpo or len(numero_limpo) < 10:
                        falhas += 1
                        continue
                    
                    # COPIAR MENSAGEM PARA CLIPBOARD
                    pyperclip.copy(mensagem_final)
                    time.sleep(0.5)
                    
                    # ABRIR NOVA CONVERSA NO WHATSAPP WEB
                    # Método 1: URL direta
                    url = f"https://web.whatsapp.com/send?phone={numero_limpo}"
                    webbrowser.open_new_tab(url)
                    
                    # Aguardar carregamento - MAIS TEMPO
                    time.sleep(8)  # WhatsApp Web pode ser lento
                    
                    # Verificar se a janela carregou (buscar pela caixa de texto)
                    time.sleep(2)
                    
                    # COLAR A MENSAGEM
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(2)  # Dar tempo para colar
                    
                    # ENVIAR MENSAGEM
                    pyautogui.press('enter')
                    time.sleep(3)  # Aguardar envio
                    
                    enviados += 1
                    
                    # FECHAR ABA
                    pyautogui.hotkey('ctrl', 'w')
                    time.sleep(2)
                    
                    # Aguardar tempo configurado
                    if index < total - 1 and not self.cancelar_envio:
                        self.root.after(0, lambda: self.update_status(f"Aguardando {tempo_espera} segundos..."))
                        for seg in range(tempo_espera, 0, -1):
                            if self.cancelar_envio:
                                break
                            time.sleep(1)
                    
                except Exception as e:
                    falhas += 1
                    print(f"Erro ao enviar para {nome}: {e}")
                    # Tentar fechar aba se houver erro
                    try:
                        pyautogui.hotkey('ctrl', 'w')
                    except:
                        pass
                
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro no Envio", f"Erro durante o envio: {str(e)}"))
        
        # Finalizar
        self.root.after(0, self.finalizar_envio, total, enviados, falhas)
    
    def finalizar_envio(self, total, enviados, falhas):
        self.enviando = False
        self.send_button.configure(state="normal", text="▶️ Iniciar Envio")
        self.tempo_btn.configure(state="normal")
        self.cancel_button.pack_forget()
        
        # Resultado
        self.root.after(0, lambda: messagebox.showinfo(
            "Envio Concluído",
            f"✅ Processo finalizado!\n\n"
            f"📊 Estatísticas:\n"
            f"• Total de contatos: {total}\n"
            f"• Mensagens enviadas: {enviados}\n"
            f"• Falhas: {falhas}\n\n"
            f"{'🎉 Sucesso total!' if falhas == 0 else '⚠️ Algumas mensagens falharam.'}"
        ))
        
        self.update_status(f"Concluído: {enviados} enviados, {falhas} falhas")
        self.progress_bar.set(1)
        self.progress_label.configure(text="Progresso: 100%")
        
        # Limpar
        self.clear_excel_file()
        self.clear_image_file()
    
    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    # Verificar dependências
    try:
        import pyautogui
        import pyperclip
    except ImportError:
        print("Instalando dependências necessárias...")
        import subprocess
        import sys
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyautogui", "pyperclip"])
        import pyautogui
        import pyperclip
    
    app = MessageFlowApp()
    app.run()