import customtkinter as ctk
import json, subprocess
from tkinter import messagebox
import os

def capturar_hwid_local():
    try:
        # Comando moderno compatível com Windows 10 e 11 (PowerShell)
        cmd = 'powershell (Get-CimInstance Win32_ComputerSystemProduct).UUID'
        
        # Captura a saída, remove espaços e decodifica
        uuid = subprocess.check_output(cmd, shell=True).decode().strip()
        
        if uuid:
            entry_hwid.delete(0, 'end')
            entry_hwid.insert(0, uuid)
        else:
            raise ValueError("UUID retornado vazio.")
            
    except Exception as e:
        # Se o PowerShell falhar, tentamos um método alternativo via Registro do Windows
        try:
            cmd_alt = 'reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Cryptography" /v MachineGuid'
            saida = subprocess.check_output(cmd_alt, shell=True).decode()
            uuid_alt = saida.split()[-1]
            entry_hwid.delete(0, 'end')
            entry_hwid.insert(0, uuid_alt)
        except:
            messagebox.showerror("Erro", "Não foi possível capturar o HWID. Tente colar manualmente.")

def gerar_arquivo():
    hwid = entry_hwid.get().strip()
    if not hwid:
        messagebox.showwarning("Atenção", "O campo HWID está vazio!")
        return
    
    # Gera o JSON com a chave correta para o seu MessageFlow ler
    dados = {"hwid_autorizado": hwid}
    
    try:
        with open("config.json", "w") as f:
            json.dump(dados, f, indent=4)
        messagebox.showinfo("Sucesso", "Arquivo config.json gerado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao salvar arquivo: {e}")

# --- INTERFACE ---
app = ctk.CTk()
app.title("Admin - Gerador de Ativação")
app.geometry("450x350")
app.resizable(False, False)

ctk.CTkLabel(app, text="Gerador de Licença MessageFlow", font=("Segoe UI", 18, "bold")).pack(pady=20)

f = ctk.CTkFrame(app)
f.pack(pady=10, padx=20, fill="x")

entry_hwid = ctk.CTkEntry(f, placeholder_text="UUID do Cliente aparecerá aqui")
entry_hwid.pack(padx=10, pady=20, fill="x")

ctk.CTkButton(app, text="1. Capturar Meu HWID", command=capturar_hwid_local, fg_color="#6C757D", height=35).pack(pady=5, padx=20, fill="x")
ctk.CTkButton(app, text="2. Gerar Licença (config.json)", command=gerar_arquivo, fg_color="#28A745", height=45, font=("Segoe UI", 12, "bold")).pack(pady=15, padx=20, fill="x")

app.mainloop()