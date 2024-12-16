import tkinter as tk
from tkinter import messagebox
import subprocess
import os

def executar_codigo():
    try:
        # Caminho para o script que será executado
        script_path = r"Caminho onde esta o codigo a ser rodado"
        
        # Verifica se o arquivo existe
        if not os.path.exists(script_path):
            messagebox.showerror("Erro", f"O arquivo não foi encontrado:\n{script_path}")
            return
        
        result = subprocess.run(
            ["python", script_path], 
            stdout=subprocess.DEVNULL,  
            stderr=subprocess.DEVNULL   
        )

        # Mostra a mensagem final com base no retorno
        if result.returncode == 0:
            messagebox.showinfo("Sucesso", "Solicitações concluídas com sucesso.")
        else:
            messagebox.showerror("Erro", "Ocorreu um erro durante a execução.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

# Criação da janela principal
janela = tk.Tk()
janela.title("Executar Código Python")
janela.geometry("400x300")
janela.configure(bg="#f0f0f0")

# Título
titulo = tk.Label(janela, text="Abertura de chamado", font=("Arial", 18, "bold"), bg="#f0f0f0")
titulo.pack(pady=20)

# Botão para executar o código
botao_executar = tk.Button(janela, text="Executar Código", font=("Arial", 14), bg="#4CAF50", fg="white", command=executar_codigo)
botao_executar.pack(pady=30)


rodape = tk.Label(janela, text="Desenvolvido por Você", font=("Arial", 10), fg="gray", bg="#f0f0f0")
rodape.pack(side="bottom", pady=10)


frame = tk.Frame(janela, bg="#f0f0f0")
frame.pack(pady=10)

janela.mainloop()
