"""
Processador em lote
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
from pathlib import Path

class BatchProcessorDialog:
    def __init__(self, parent, engine):
        self.parent = parent
        self.engine = engine
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Processamento em Lote - DataCraft")
        self.dialog.geometry("700x600")
        
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.files_list = []
        
        self.criar_interface()
        self.dialog.transient(parent)
        self.dialog.grab_set()
    
    def criar_interface(self):
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Diretório de entrada
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="📁 Diretório de Entrada:").pack(side=tk.LEFT)
        ttk.Entry(input_frame, textvariable=self.input_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(input_frame, text="Buscar", command=self.selecionar_input).pack(side=tk.RIGHT)
        
        # Diretório de saída
        output_frame = ttk.Frame(main_frame)
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="💾 Diretório de Saída:").pack(side=tk.LEFT)
        ttk.Entry(output_frame, textvariable=self.output_dir).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(output_frame, text="Buscar", command=self.selecionar_output).pack(side=tk.RIGHT)
        
        # Lista de arquivos
        ttk.Label(main_frame, text="📄 Arquivos encontrados:").pack(anchor=tk.W, pady=(10, 5))
        
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.files_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.files_listbox.yview)
        
        # Progresso
        self.progress_var = tk.StringVar(value="Aguardando...")
        ttk.Label(main_frame, textvariable=self.progress_var).pack(pady=(10, 5))
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        # Botões
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        self.btn_processar = ttk.Button(btn_frame, text="🚀 Processar Todos", command=self.processar)
        self.btn_processar.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="Fechar", command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def selecionar_input(self):
        diretorio = filedialog.askdirectory(title="Selecione o diretório com os arquivos")
        if diretorio:
            self.input_dir.set(diretorio)
            self.carregar_arquivos()
    
    def selecionar_output(self):
        diretorio = filedialog.askdirectory(title="Selecione o diretório de saída")
        if diretorio:
            self.output_dir.set(diretorio)
    
    def carregar_arquivos(self):
        self.files_listbox.delete(0, tk.END)
        self.files_list = []
        
        extensoes = ['.xml', '.json', '.csv', '.xlsx', '.yaml', '.yml']
        
        for arquivo in Path(self.input_dir.get()).iterdir():
            if arquivo.suffix.lower() in extensoes:
                self.files_list.append(str(arquivo))
                self.files_listbox.insert(tk.END, arquivo.name)
        
        self.progress_var.set(f"Encontrados {len(self.files_list)} arquivos")
    
    def processar(self):
        if not self.files_list:
            messagebox.showwarning("Aviso", "Nenhum arquivo para processar!")
            return
        
        if not self.output_dir.get():
            messagebox.showwarning("Aviso", "Selecione o diretório de saída!")
            return
        
        self.btn_processar.config(state='disabled')
        thread = threading.Thread(target=self.executar_batch)
        thread.daemon = True
        thread.start()
    
    def executar_batch(self):
        total = len(self.files_list)
        sucessos = 0
        erros = 0
        
        for i, arquivo in enumerate(self.files_list):
            nome = Path(arquivo).stem
            saida = os.path.join(self.output_dir.get(), f"{nome}_convertido.xlsx")
            
            try:
                self.progress_var.set(f"Processando: {Path(arquivo).name} ({i+1}/{total})")
                self.progress_bar['value'] = (i + 1) / total * 100
                
                self.engine.convert(arquivo, saida)
                sucessos += 1
                
            except Exception as e:
                erros += 1
                print(f"Erro em {arquivo}: {e}")
        
        messagebox.showinfo("Concluído!", 
            f"Processamento em lote finalizado!\n\n"
            f"✅ Sucessos: {sucessos}\n"
            f"❌ Erros: {erros}\n"
            f"📁 Total: {total}")
        
        self.btn_processar.config(state='normal')
        self.progress_var.set("Processamento concluído!")