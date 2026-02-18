import customtkinter as ctk
import json, subprocess
from tkinter import messagebox

def capturar_hwid_local():
    try:
        cmd = 'wmic csproduct get uuid'
        uuid = str(subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip())
        entry_hwid.delete(0, 'end')
        entry_hwid.insert(0, uuid)
    except:
        messagebox.showerror("Erro", "Não foi possível capturar o HWID automaticamente.")

def gerar_arquivo():
    hwid = entry_hwid.get().strip()
    if not hwid:
        messagebox.showwarning("Atenção", "O campo HWID está vazio!")
        return
    
    with open("config.json", "w") as f:
        json.dump({"hwid_autorizado": hwid}, f, indent=4)
    messagebox.showinfo("Sucesso", "Arquivo config.json gerado com sucesso!")

app = ctk.CTk()
app.title("Admin - Gerador de Ativação")
app.geometry("450x300")

ctk.CTkLabel(app, text="Gerador de Licença MessageFlow", font=("Segoe UI", 18, "bold")).pack(pady=20)

f = ctk.CTkFrame(app)
f.pack(pady=10, padx=20, fill="x")

entry_hwid = ctk.CTkEntry(f, placeholder_text="Cole aqui o HWID do cliente")
entry_hwid.pack(side="left", fill="x", expand=True, padx=5, pady=10)

ctk.CTkButton(app, text="Capturar Meu HWID", command=capturar_hwid_local, fg_color="#6C757D").pack(pady=5)
ctk.CTkButton(app, text="Gerar Licença (config.json)", command=gerar_arquivo, fg_color="#28A745").pack(pady=10)

app.mainloop()