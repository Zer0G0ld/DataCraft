"""
Diálogo para mapeamento de colunas
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class ColumnMapperDialog:
    def __init__(self, parent, dataframe):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Mapeamento de Colunas - DataCraft")
        self.dialog.geometry("600x500")
        self.df = dataframe
        self.mapping = {}
        
        self.criar_interface()
        self.dialog.transient(parent)
        self.dialog.grab_set()
    
    def criar_interface(self):
        # Frame principal
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Instruções
        ttk.Label(main_frame, text="Mapeie as colunas do arquivo original para os nomes desejados:",
                 font=('Segoe UI', 10)).pack(pady=(0, 10))
        
        # Frame para scroll
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(canvas_frame)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Cabeçalho
        ttk.Label(scrollable_frame, text="Coluna Original", font=('Segoe UI', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5)
        ttk.Label(scrollable_frame, text="→", font=('Segoe UI', 10)).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(scrollable_frame, text="Novo Nome", font=('Segoe UI', 10, 'bold')).grid(row=0, column=2, padx=5, pady=5)
        
        # Entradas para cada coluna
        self.entries = {}
        for i, col in enumerate(self.df.columns, 1):
            ttk.Label(scrollable_frame, text=col).grid(row=i, column=0, padx=5, pady=2, sticky='w')
            ttk.Label(scrollable_frame, text="→").grid(row=i, column=1, padx=5, pady=2)
            
            entry = ttk.Entry(scrollable_frame, width=30)
            entry.insert(0, col)
            entry.grid(row=i, column=2, padx=5, pady=2)
            self.entries[col] = entry
        
        # Botões
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Aplicar Mapeamento", command=self.aplicar).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancelar", command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def aplicar(self):
        self.mapping = {col: entry.get() for col, entry in self.entries.items() if entry.get() != col}
        self.dialog.destroy()