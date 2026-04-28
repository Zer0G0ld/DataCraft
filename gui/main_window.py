"""
Interface principal do DataCraft - Versão refatorada para usar o novo motor
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
from datetime import datetime
import pandas as pd

# Importa o motor universal
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from core.engine import ConversionEngine
from gui.mapper_dialog import ColumnMapperDialog
from gui.batch_processor import BatchProcessorDialog

class DataCraftGUI:
    """Aplicativo principal com interface gráfica"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("DataCraft 🏆 - Transforme dados em ouro - Universal Converter")
        self.root.geometry("900x800")
        self.root.resizable(True, True)
        
        # Inicializa o motor universal
        self.engine = ConversionEngine()
        self.engine.set_progress_callback(self.update_progress)
        
        # Variáveis
        self.arquivo_entrada = tk.StringVar()
        self.arquivo_saida = tk.StringVar()
        self.formato_entrada = tk.StringVar(value="auto")
        self.formato_saida = tk.StringVar(value="auto")
        self.status_var = tk.StringVar(value="✨ Pronto para começar!")
        
        # Configurações
        self.configurar_estilo()
        self.criar_interface()
        self.centralizar_janela()
        
        # Log inicial
        self.log("🚀 DataCraft Universal iniciado!", 'success')
        self.log(f"📥 Formatos suportados (entrada): {', '.join(self.engine.get_supported_inputs())}", 'info')
        self.log(f"📤 Formatos suportados (saída): {', '.join(self.engine.get_supported_outputs())}", 'info')
    
    def configurar_estilo(self):
        """Configura o estilo visual"""
        self.bg_color = "#f8f9fa"
        self.primary_color = "#2c3e50"
        self.success_color = "#27ae60"
        self.error_color = "#e74c3c"
        
        self.root.configure(bg=self.bg_color)
        
        # Tenta carregar ícone
        try:
            if os.path.exists('voto.ico'):
                self.root.iconbitmap('voto.ico')
        except:
            pass
    
    def centralizar_janela(self):
        """Centraliza a janela na tela"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def criar_interface(self):
        """Cria todos os elementos da interface"""
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.criar_titulo(main_frame)
        self.criar_frame_arquivos(main_frame)
        self.criar_frame_formatos(main_frame)
        self.criar_opcoes(main_frame)
        self.criar_botoes_acao(main_frame)
        self.criar_status_progresso(main_frame)
        self.criar_area_log(main_frame)
    
    def criar_titulo(self, parent):
        """Cria o título"""
        titulo = tk.Label(
            parent,
            text="🏆 DataCraft - Transforme dados em ouro",
            font=('Segoe UI', 20, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color
        )
        titulo.pack(pady=(0, 5))
        
        subtitulo = tk.Label(
            parent,
            text="Conversor Universal: Qualquer formato → Qualquer formato",
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg='#666'
        )
        subtitulo.pack(pady=(0, 20))
    
    def criar_frame_arquivos(self, parent):
        """Frame de seleção de arquivos"""
        frame = tk.LabelFrame(parent, text="📁 Arquivos", font=('Segoe UI', 11, 'bold'),
                              bg=self.bg_color, padx=10, pady=10)
        frame.pack(fill=tk.X, pady=(0, 10))
        
        # Entrada
        entrada_frame = tk.Frame(frame, bg=self.bg_color)
        entrada_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(entrada_frame, text="Entrada:", width=10, anchor='w',
                bg=self.bg_color, font=('Segoe UI', 10)).pack(side=tk.LEFT)
        
        tk.Entry(entrada_frame, textvariable=self.arquivo_entrada,
                font=('Segoe UI', 10), bg='white').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        tk.Button(entrada_frame, text="📂 Buscar", command=self.selecionar_entrada,
                 bg=self.primary_color, fg='white', cursor='hand2').pack(side=tk.RIGHT)
        
        # Saída
        saida_frame = tk.Frame(frame, bg=self.bg_color)
        saida_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(saida_frame, text="Saída:", width=10, anchor='w',
                bg=self.bg_color, font=('Segoe UI', 10)).pack(side=tk.LEFT)
        
        tk.Entry(saida_frame, textvariable=self.arquivo_saida,
                font=('Segoe UI', 10), bg='white').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        tk.Button(saida_frame, text="💾 Salvar", command=self.selecionar_saida,
                 bg=self.primary_color, fg='white', cursor='hand2').pack(side=tk.RIGHT)
    
    def criar_frame_formatos(self, parent):
        """Frame de seleção de formatos"""
        frame = tk.LabelFrame(parent, text="🔄 Formatos", font=('Segoe UI', 11, 'bold'),
                              bg=self.bg_color, padx=10, pady=10)
        frame.pack(fill=tk.X, pady=(0, 10))
        
        formatos_frame = tk.Frame(frame, bg=self.bg_color)
        formatos_frame.pack()
        
        # Formato entrada
        tk.Label(formatos_frame, text="Formato entrada:", bg=self.bg_color,
                font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=5)
        
        entrada_opts = ['auto'] + self.engine.get_supported_inputs()
        self.formato_entrada_combo = ttk.Combobox(formatos_frame, textvariable=self.formato_entrada,
                                                   values=entrada_opts, width=12, state='readonly')
        self.formato_entrada_combo.pack(side=tk.LEFT, padx=5)
        
        tk.Label(formatos_frame, text="➡️", bg=self.bg_color,
                font=('Segoe UI', 12)).pack(side=tk.LEFT, padx=10)
        
        # Formato saída
        tk.Label(formatos_frame, text="Formato saída:", bg=self.bg_color,
                font=('Segoe UI', 10)).pack(side=tk.LEFT, padx=5)
        
        saida_opts = ['auto'] + self.engine.get_supported_outputs()
        self.formato_saida_combo = ttk.Combobox(formatos_frame, textvariable=self.formato_saida,
                                                 values=saida_opts, width=12, state='readonly')
        self.formato_saida_combo.pack(side=tk.LEFT, padx=5)
    
    def criar_opcoes(self, parent):
        """Opções de conversão"""
        frame = tk.LabelFrame(parent, text="⚙️ Opções", font=('Segoe UI', 11, 'bold'),
                              bg=self.bg_color, padx=10, pady=10)
        frame.pack(fill=tk.X, pady=(0, 10))
        
        self.flatten_var = tk.BooleanVar(value=True)
        tk.Checkbutton(frame, text="📊 Achatar estruturas aninhadas (expandir listas/dicionários)",
                      variable=self.flatten_var, bg=self.bg_color).pack(anchor=tk.W)
        
        self.auto_abrir_var = tk.BooleanVar(value=True)
        tk.Checkbutton(frame, text="🚀 Abrir arquivo após conversão",
                      variable=self.auto_abrir_var, bg=self.bg_color).pack(anchor=tk.W)
        
        self.summary_var = tk.BooleanVar(value=True)
        tk.Checkbutton(frame, text="📈 Gerar aba resumida (para Excel)",
                      variable=self.summary_var, bg=self.bg_color).pack(anchor=tk.W)
    
    def criar_botoes_acao(self, parent):
        """Botões de ação"""
        btn_frame = tk.Frame(parent, bg=self.bg_color)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.btn_converter = tk.Button(btn_frame, text="🔄 CONVERTER AGORA",
                                       command=self.converter, bg=self.success_color,
                                       fg='white', font=('Segoe UI', 12, 'bold'),
                                       height=2, cursor='hand2')
        self.btn_converter.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        tk.Button(btn_frame, text="🗺️ Mapear Colunas", command=self.abrir_mapper,
                 bg='#3498db', fg='white', font=('Segoe UI', 10), cursor='hand2').pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="📦 Processamento em Lote", command=self.abrir_batch,
                 bg='#9b59b6', fg='white', font=('Segoe UI', 10), cursor='hand2').pack(side=tk.LEFT, padx=5)
    
    def criar_status_progresso(self, parent):
        """Barra de status e progresso"""
        status_frame = tk.Frame(parent, bg=self.bg_color)
        status_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Label(status_frame, textvariable=self.status_var, bg=self.bg_color,
                font=('Segoe UI', 9)).pack(side=tk.LEFT)
        
        self.progress_bar = ttk.Progressbar(parent, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
    
    def criar_area_log(self, parent):
        """Área de log"""
        log_frame = tk.LabelFrame(parent, text="📝 Log de Conversão", font=('Segoe UI', 11, 'bold'),
                                  bg=self.bg_color, padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        text_frame = tk.Frame(log_frame, bg=self.bg_color)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(text_frame, height=15, font=('Consolas', 9),
                                bg='#1e1e1e', fg='#d4d4d4', wrap=tk.WORD)
        scrollbar = tk.Scrollbar(text_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botões do log
        btn_frame = tk.Frame(log_frame, bg=self.bg_color)
        btn_frame.pack(pady=(5, 0))
        
        tk.Button(btn_frame, text="🗑️ Limpar", command=self.limpar_log,
                 bg='#95a5a6', fg='white', cursor='hand2').pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="📋 Copiar", command=self.copiar_log,
                 bg='#95a5a6', fg='white', cursor='hand2').pack(side=tk.LEFT, padx=5)
    
    def selecionar_entrada(self):
        """Seleciona arquivo de entrada"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo",
            filetypes=[
                ("Todos os formatos suportados", "*.xml;*.json;*.csv;*.xlsx;*.yaml;*.yml"),
                ("XML", "*.xml"), ("JSON", "*.json"), ("CSV", "*.csv"),
                ("Excel", "*.xlsx"), ("YAML", "*.yaml;*.yml"), ("Todos", "*.*")
            ]
        )
        if arquivo:
            self.arquivo_entrada.set(arquivo)
            # Sugere nome de saída
            nome_base = os.path.splitext(arquivo)[0]
            self.arquivo_saida.set(f"{nome_base}_convertido.xlsx")
            self.log(f"Arquivo selecionado: {arquivo}", 'success')
    
    def selecionar_saida(self):
        """Seleciona arquivo de saída"""
        formato = self.formato_saida.get()
        if formato == 'auto':
            formato = 'xlsx'
        
        extensoes = {
            'excel': [('Excel', '*.xlsx')],
            'csv': [('CSV', '*.csv')],
            'json': [('JSON', '*.json')],
            'html': [('HTML', '*.html')],
            'markdown': [('Markdown', '*.md')],
            'sql': [('SQL', '*.sql')]
        }
        
        filetypes = extensoes.get(formato, [('Excel', '*.xlsx')])
        
        arquivo = filedialog.asksaveasfilename(
            title="Salvar arquivo como",
            defaultextension=f".{formato if formato != 'excel' else 'xlsx'}",
            filetypes=filetypes
        )
        if arquivo:
            self.arquivo_saida.set(arquivo)
            self.log(f"Arquivo de saída: {arquivo}", 'info')
    
    def abrir_mapper(self):
        """Abre diálogo de mapeamento de colunas"""
        if self.engine.last_dataframe is not None:
            dialog = ColumnMapperDialog(self.root, self.engine.last_dataframe)
            if dialog.mapping:
                self.log(f"🗺️ Mapeamento aplicado: {dialog.mapping}", 'info')
        else:
            messagebox.showinfo("Info", "Converta um arquivo primeiro para usar o mapeador!")
    
    def abrir_batch(self):
        """Abre diálogo de processamento em lote"""
        dialog = BatchProcessorDialog(self.root, self.engine)
    
    def converter(self):
        """Inicia conversão"""
        if not self.arquivo_entrada.get():
            messagebox.showerror("Erro", "Selecione um arquivo de entrada!")
            return
        
        if not self.arquivo_saida.get():
            messagebox.showerror("Erro", "Defina o arquivo de saída!")
            return
        
        self.btn_converter.config(state='disabled', text='🔄 CONVERTENDO...')
        self.progress_bar['value'] = 0
        self.limpar_log()
        
        thread = threading.Thread(target=self.executar_conversao)
        thread.daemon = True
        thread.start()
    
    def executar_conversao(self):
        """Executa conversão em thread separada"""
        try:
            formato_entrada = None if self.formato_entrada.get() == 'auto' else self.formato_entrada.get()
            formato_saida = None if self.formato_saida.get() == 'auto' else self.formato_saida.get()
            
            options = {
                'flatten': self.flatten_var.get(),
                'writer_options': {
                    'create_summary': self.summary_var.get(),
                    'auto_width': True
                }
            }
            
            metadata = self.engine.convert(
                input_path=self.arquivo_entrada.get(),
                output_path=self.arquivo_saida.get(),
                input_format=formato_entrada,
                output_format=formato_saida,
                options=options
            )
            
            self.log("=" * 50, 'success')
            self.log("✨ CONVERSÃO CONCLUÍDA COM SUCESSO!", 'success')
            self.log(f"📊 {metadata['rows']} registros × {metadata['columns']} colunas")
            self.log(f"⏱️ Tempo: {metadata['duration_seconds']:.1f} segundos")
            self.log(f"📁 Saída: {metadata['output_path']}")
            
            if self.auto_abrir_var.get():
                self.abrir_arquivo(metadata['output_path'])
            
            messagebox.showinfo("Sucesso!", 
                f"Conversão concluída!\n\n"
                f"📊 {metadata['rows']} registros\n"
                f"📋 {metadata['columns']} colunas\n"
                f"⏱️ {metadata['duration_seconds']:.1f} segundos")
            
        except Exception as e:
            self.log(f"❌ ERRO: {str(e)}", 'error')
            messagebox.showerror("Erro", f"Falha na conversão:\n{str(e)}")
        
        finally:
            self.btn_converter.config(state='normal', text='🔄 CONVERTER AGORA')
    
    def update_progress(self, message: str, percentage: int = None):
        """Atualiza barra de progresso"""
        self.log(message, 'info')
        if percentage is not None:
            self.progress_bar['value'] = percentage
            self.root.update_idletasks()
    
    def abrir_arquivo(self, caminho):
        """Abre arquivo com programa padrão"""
        try:
            if os.name == 'nt':
                os.startfile(caminho)
            else:
                import subprocess
                subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', caminho])
        except Exception as e:
            self.log(f"Não foi possível abrir: {e}", 'warning')
    
    def log(self, mensagem, tipo='info'):
        """Adiciona mensagem ao log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        emojis = {'error': '❌', 'success': '✅', 'warning': '⚠️', 'info': 'ℹ️'}
        prefixos = {'error': 'ERRO', 'success': 'OK', 'warning': 'AVISO', 'info': 'INFO'}
        
        self.log_text.insert(tk.END, f"[{timestamp}] {emojis.get(tipo, 'ℹ️')} {prefixos.get(tipo, 'INFO')}: {mensagem}\n", tipo)
        self.log_text.see(tk.END)
        
        # Configura cores
        self.log_text.tag_config('error', foreground='#e74c3c')
        self.log_text.tag_config('success', foreground='#27ae60')
        self.log_text.tag_config('warning', foreground='#f39c12')
        self.log_text.tag_config('info', foreground='#3498db')
        
        self.root.update_idletasks()
    
    def limpar_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def copiar_log(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.log_text.get(1.0, tk.END))

def main():
    root = tk.Tk()
    app = DataCraftGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()