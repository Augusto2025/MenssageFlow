import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import pandas as pd
import time
import os
import pywhatkit as kit
import webbrowser
import re
import traceback
import sys
import json
import hashlib
import uuid
import tempfile
from contextlib import contextmanager

# Configurações de Tempo
SEND_WAIT_TIME_IMAGE = 35
SEND_CLOSE_TIME_IMAGE = 6
SEND_WAIT_TIME_TEXT = 20
SEND_CLOSE_TIME_TEXT = 4
SEND_BETWEEN = 5
ATTEMPT_SLEEP_MULTIPLIER = 5

def iniciar_sessao():
    webbrowser.open('https://web.whatsapp.com/')
    messagebox.showinfo('Instrução', '⚠️ Abra o WhatsApp Web e faça login.\nDepois feche a janela e clique em OK para continuar.')

def escolher_excel():
    arquivo = filedialog.askopenfilename(title='Selecione o arquivo Excel', filetypes=[('Planilhas Excel', '*.xlsx')])
    if arquivo:
        caminho_excel.set(arquivo)
        atualizar_status_arquivo_imagem()

def escolher_imagem():
    imagem = filedialog.askopenfilename(title='Selecione a imagem (opcional)', filetypes=[('Imagens', '*.jpg;*.png;*.jpeg')])
    if imagem:
        caminho_imagem.set(imagem)
    else:
        caminho_imagem.set('')
    atualizar_status_arquivo_imagem()

def atualizar_status_arquivo_imagem():
    excel = caminho_excel.get()
    imagem = caminho_imagem.get()
    
    status = "Aguardando seleção de arquivo..."
    if excel:
        status = f"Excel: {os.path.basename(excel)}"
    if imagem:
        status += f" | Imagem: {os.path.basename(imagem)}"
    
    label_arquivo_imagem.config(text=status, fg="green" if excel else "gray")

def formatar_telefone(raw):
    if pd.isna(raw):
        return None
    try:
        s = str(raw).strip()
        digits = re.sub(r'\D', '', s)
        if not digits:
            return None
        digits = digits.lstrip('0')
        if len(digits) in (10, 11):
            digits = '55' + digits
        if 12 <= len(digits) <= 15:
            return '+' + digits
        return None
    except:
        return None

def set_controls_state(enabled):
    state = 'normal' if enabled else 'disabled'
    btn_iniciar_sessao.config(state=state)
    btn_selecionar_excel.config(state=state)
    btn_selecionar_imagem.config(state=state)
    btn_iniciar_envio.config(state=state)
    if not enabled:
        label_status.config(text='Enviando... por favor aguarde', fg='blue')
    root.update_idletasks()

def _cleanup_pywhatkit_db():
    try:
        db_path = os.path.join(tempfile.gettempdir(), 'PyWhatKit_DB.txt')
        if os.path.exists(db_path):
            os.remove(db_path)
    except:
        pass

def safe_send_image(phone, imagem, caption, max_attempts):
    for attempt in range(1, max_attempts + 1):
        try:
            kit.sendwhats_image(phone, imagem, caption=caption, wait_time=SEND_WAIT_TIME_IMAGE, tab_close=True, close_time=SEND_CLOSE_TIME_IMAGE)
            _cleanup_pywhatkit_db()
            return (True, None)
        except Exception as e:
            _cleanup_pywhatkit_db()
            time.sleep(ATTEMPT_SLEEP_MULTIPLIER * attempt)
            last_exception = e
    return (False, last_exception)

def safe_send_text(phone, message, max_attempts):
    for attempt in range(1, max_attempts + 1):
        try:
            kit.sendwhatmsg_instantly(phone, message, wait_time=SEND_WAIT_TIME_TEXT, tab_close=True, close_time=SEND_CLOSE_TIME_TEXT)
            _cleanup_pywhatkit_db()
            return (True, None)
        except Exception as e:
            _cleanup_pywhatkit_db()
            time.sleep(ATTEMPT_SLEEP_MULTIPLIER * attempt)
            last_exception = e
    return (False, last_exception)

def extrair_primeiro_nome(nome):
    try:
        if not nome: return ''
        s = str(nome).strip()
        partes = re.findall(r'[A-Za-zÀ-ÖØ-öø-ÿ]+', s, flags=re.UNICODE)
        return partes[0].capitalize() if partes else s.split()[0].capitalize()
    except:
        return ''

def enviar_mensagens():
    arquivo = caminho_excel.get()
    imagem = caminho_imagem.get()
    mensagem_base = campo_mensagem.get('1.0', tk.END).strip()

    if not arquivo or not mensagem_base:
        messagebox.showwarning('Atenção', 'Verifique se o Excel e a Mensagem foram preenchidos!')
        set_controls_state(True)
        return

    try:
        df = pd.read_excel(arquivo)
        total = len(df)
        progresso['maximum'] = total

        for index, row in df.iterrows():
            nome_original = str(row.get('Nome', 'Cliente'))
            primeiro_nome = extrair_primeiro_nome(nome_original)
            telefone = formatar_telefone(row.get('Telefone'))
            
            if not telefone: continue
            
            msg_personalizada = mensagem_base.replace('{nome}', primeiro_nome)
            label_status.config(text=f"Enviando para {primeiro_nome} ({index+1}/{total})...")
            
            if imagem and os.path.exists(imagem):
                safe_send_image(telefone, imagem, msg_personalizada, 2)
            else:
                safe_send_text(telefone, msg_personalizada, 2)
            
            progresso['value'] = index + 1
            root.update_idletasks()
            time.sleep(SEND_BETWEEN)
            
        messagebox.showinfo('Sucesso', 'Envio finalizado!')
    except Exception as e:
        messagebox.showerror('Erro', f'Falha ao processar: {str(e)}')
    finally:
        set_controls_state(True)
        label_status.config(text='Concluído', fg='green')

def iniciar_envio_thread():
    set_controls_state(False)
    threading.Thread(target=enviar_mensagens, daemon=True).start()

def limpar_excel():
    caminho_excel.set('')
    atualizar_status_arquivo_imagem()

def limpar_imagem():
    caminho_imagem.set('')
    atualizar_status_arquivo_imagem()

def _app_path(filename):
    base = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)

def obter_hwid():
    nome = os.getenv('COMPUTERNAME') or os.getenv('HOSTNAME') or 'PC-USER'
    mac = str(uuid.getnode())
    return hashlib.sha256((nome + mac).encode()).hexdigest()

def check_local_binding():
    # Se quiser ignorar a trava de segurança para testes, mude para: return True
    current = obter_hwid()
    bind_path = _app_path('bind.json')
    try:
        if os.path.exists(bind_path):
            with open(bind_path, 'r') as f:
                return json.load(f).get('hwid') == current
        else:
            with open(bind_path, 'w') as f:
                json.dump({'hwid': current}, f)
            return True
    except:
        return False

# Interface Gráfica
root = tk.Tk()
root.title('MenssageFlow')
root.geometry('600x720')
root.resizable(False, False)
root.configure(bg='#f4f4f4')

caminho_excel = tk.StringVar()
caminho_imagem = tk.StringVar()

tk.Label(root, text='⚠️ É necessário estar conectado ao WhatsApp Web.', bg='#fff3cd', fg='#856404', font=('Segoe UI', 10, 'bold'), wraplength=550, relief='solid', bd=1).pack(pady=10, fill='x')
tk.Label(root, text='💬 MenssageFlow', font=('Segoe UI', 18, 'bold'), bg='#f4f4f4', fg='#004aad').pack(pady=10)

frame = tk.Frame(root, bg='#f4f4f4')
frame.pack(pady=5)

btn_iniciar_sessao = tk.Button(frame, text='Iniciar Sessão', command=iniciar_sessao, bg='#004aad', fg='white', width=44, height=2)
btn_iniciar_sessao.grid(row=0, column=0, columnspan=2, pady=5)

btn_selecionar_excel = tk.Button(frame, text='Selecionar Excel', command=escolher_excel, bg='#004aad', fg='white', width=28, height=2)
btn_selecionar_excel.grid(row=1, column=0, padx=5, pady=5)

btn_limpar_excel = tk.Button(frame, text='❌ Limpar Excel', command=limpar_excel, bg='#b23b3b', fg='white', width=14, height=2)
btn_limpar_excel.grid(row=1, column=1, padx=5, pady=5)

btn_selecionar_imagem = tk.Button(frame, text='Selecionar Imagem', command=escolher_imagem, bg='#004aad', fg='white', width=28, height=2)
btn_selecionar_imagem.grid(row=2, column=0, padx=5, pady=5)

btn_limpar_imagem = tk.Button(frame, text='❌ Limpar Imagem', command=limpar_imagem, bg='#b23b3b', fg='white', width=14, height=2)
btn_limpar_imagem.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text='Mensagem (use {nome} para personalizar):', bg='#f4f4f4').pack(pady=5)
campo_mensagem = tk.Text(root, height=8, width=65)
campo_mensagem.pack(pady=5)

btn_iniciar_envio = tk.Button(root, text='🚀 Iniciar Envio', command=iniciar_envio_thread, bg='#00a859', fg='white', width=25, height=2)
btn_iniciar_envio.pack(pady=20)

progresso = ttk.Progressbar(root, orient='horizontal', length=450, mode='determinate')
progresso.pack(pady=10)

label_arquivo_imagem = tk.Label(root, text='Aguardando seleção...', bg='#f4f4f4', fg='gray')
label_arquivo_imagem.pack()

label_status = tk.Label(root, text='', bg='#f4f4f4')
label_status.pack(pady=10)

# Verificação de Segurança ao Iniciar
if not check_local_binding():
    messagebox.showerror('Bloqueado', 'Arquivo não autorizado nesta máquina.')
    sys.exit(1)

root.mainloop()