"""
DataCraft - Conversor Inteligente de XML para Excel
Desenvolvido por Zer0G0ld 🏆
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import xml.etree.ElementTree as ET
import pandas as pd
import os
import re
import sys
from datetime import datetime

class DataCraft:
    """Aplicativo principal para conversão de XML para Excel"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("DataCraft 🏆 - Transforme dados em ouro")
        self.root.geometry("850x750")
        self.root.resizable(True, True)
        
        # Configurar ícone
        self.configurar_icone()
        
        # Configurar estilo
        self.configurar_estilo()
        
        # Variáveis
        self.arquivo_entrada = tk.StringVar()
        self.arquivo_saida = tk.StringVar()
        self.status_var = tk.StringVar(value="✨ Pronto para começar!")
        
        # Criar interface
        self.criar_interface()
        
        # Centralizar janela
        self.centralizar_janela()
        
        # Log de inicialização
        self.log("DataCraft iniciado com sucesso!", 'success')
        self.log("Selecione um arquivo XML para começar...", 'info')
    
    def configurar_icone(self):
        """Configura o ícone da aplicação"""
        try:
            # Tenta diferentes formas de carregar o ícone
            icon_paths = []
            
            # Caminho direto
            if os.path.exists('voto.ico'):
                icon_paths.append('voto.ico')
            
            # Caminho para executável
            if getattr(sys, 'frozen', False):
                application_path = os.path.dirname(sys.executable)
                icon_path = os.path.join(application_path, 'voto.ico')
                if os.path.exists(icon_path):
                    icon_paths.append(icon_path)
            
            # Tenta carregar o ícone
            for path in icon_paths:
                try:
                    self.root.iconbitmap(path)
                    break
                except:
                    continue
                    
        except Exception as e:
            print(f"Ícone não carregado: {e}")
    
    def configurar_estilo(self):
        """Configura o estilo visual do aplicativo"""
        style = ttk.Style()
        style.theme_use('vista' if os.name == 'nt' else 'default')
        
        # Cores personalizadas
        self.bg_color = "#f8f9fa"
        self.primary_color = "#2c3e50"
        self.success_color = "#27ae60"
        self.error_color = "#e74c3c"
        self.warning_color = "#f39c12"
        self.info_color = "#3498db"
        
        self.root.configure(bg=self.bg_color)
        
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
        # Frame principal
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Título
        self.criar_titulo(main_frame)
        
        # Frame de arquivos
        self.criar_frame_arquivos(main_frame)
        
        # Opções
        self.criar_opcoes(main_frame)
        
        # Botão Converter
        self.criar_botao_converter(main_frame)
        
        # Status e progresso
        self.criar_status_progresso(main_frame)
        
        # Área de log
        self.criar_area_log(main_frame)
    
    def criar_titulo(self, parent):
        """Cria o título do aplicativo"""
        titulo_frame = tk.Frame(parent, bg=self.bg_color)
        titulo_frame.pack(fill=tk.X, pady=(0, 20))
        
        titulo = tk.Label(
            titulo_frame,
            text="🏆 DataCraft - Transforme dados em ouro",
            font=('Segoe UI', 20, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color
        )
        titulo.pack()
        
        subtitulo = tk.Label(
            titulo_frame,
            text="Converta arquivos XML para Excel de forma rápida e inteligente",
            font=('Segoe UI', 10),
            bg=self.bg_color,
            fg='#666'
        )
        subtitulo.pack()
    
    def criar_frame_arquivos(self, parent):
        """Cria o frame de seleção de arquivos"""
        arquivos_frame = tk.LabelFrame(
            parent,
            text="📁 Arquivos",
            font=('Segoe UI', 11, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color,
            padx=10,
            pady=10
        )
        arquivos_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Arquivo de entrada
        entrada_frame = tk.Frame(arquivos_frame, bg=self.bg_color)
        entrada_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            entrada_frame,
            text="Arquivo XML:",
            width=12,
            anchor='w',
            font=('Segoe UI', 10),
            bg=self.bg_color
        ).pack(side=tk.LEFT)
        
        self.entrada_entry = tk.Entry(
            entrada_frame,
            textvariable=self.arquivo_entrada,
            font=('Segoe UI', 10),
            bg='white'
        )
        self.entrada_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        
        tk.Button(
            entrada_frame,
            text="📂 Buscar",
            command=self.selecionar_entrada,
            bg=self.primary_color,
            fg='white',
            font=('Segoe UI', 9),
            cursor='hand2',
            padx=10
        ).pack(side=tk.RIGHT)
        
        # Arquivo de saída
        saida_frame = tk.Frame(arquivos_frame, bg=self.bg_color)
        saida_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            saida_frame,
            text="Arquivo Excel:",
            width=12,
            anchor='w',
            font=('Segoe UI', 10),
            bg=self.bg_color
        ).pack(side=tk.LEFT)
        
        self.saida_entry = tk.Entry(
            saida_frame,
            textvariable=self.arquivo_saida,
            font=('Segoe UI', 10),
            bg='white'
        )
        self.saida_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        
        tk.Button(
            saida_frame,
            text="💾 Salvar",
            command=self.selecionar_saida,
            bg=self.primary_color,
            fg='white',
            font=('Segoe UI', 9),
            cursor='hand2',
            padx=10
        ).pack(side=tk.RIGHT)
    
    def criar_opcoes(self, parent):
        """Cria as opções avançadas"""
        opcoes_frame = tk.LabelFrame(
            parent,
            text="⚙️ Opções",
            font=('Segoe UI', 11, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color,
            padx=10,
            pady=10
        )
        opcoes_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.gerar_resumo = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opcoes_frame,
            text="📊 Gerar aba resumida com principais campos",
            variable=self.gerar_resumo,
            bg=self.bg_color,
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W)
        
        self.auto_abrir = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opcoes_frame,
            text="🚀 Abrir arquivo após conversão",
            variable=self.auto_abrir,
            bg=self.bg_color,
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W)
        
        self.excluir_vazios = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opcoes_frame,
            text="🧹 Excluir colunas totalmente vazias",
            variable=self.excluir_vazios,
            bg=self.bg_color,
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W)
    
    def criar_botao_converter(self, parent):
        """Cria o botão de conversão"""
        self.btn_converter = tk.Button(
            parent,
            text="🔄 CONVERTER AGORA",
            command=self.converter,
            bg=self.success_color,
            fg='white',
            font=('Segoe UI', 13, 'bold'),
            height=2,
            cursor='hand2',
            relief=tk.FLAT
        )
        self.btn_converter.pack(fill=tk.X, pady=(0, 15))
    
    def criar_status_progresso(self, parent):
        """Cria a barra de status e progresso"""
        status_frame = tk.Frame(parent, bg=self.bg_color)
        status_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=('Segoe UI', 9),
            bg=self.bg_color,
            fg=self.primary_color
        )
        self.status_label.pack(side=tk.LEFT)
        
        self.progress = ttk.Progressbar(parent, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(0, 10))
    
    def criar_area_log(self, parent):
        """Cria a área de log"""
        log_frame = tk.LabelFrame(
            parent,
            text="📝 Log de Conversão",
            font=('Segoe UI', 11, 'bold'),
            bg=self.bg_color,
            fg=self.primary_color,
            padx=10,
            pady=10
        )
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # Text widget com scrollbar
        text_frame = tk.Frame(log_frame, bg=self.bg_color)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(
            text_frame,
            height=18,
            font=('Consolas', 9),
            bg='#1e1e1e',
            fg='#d4d4d4',
            wrap=tk.WORD,
            relief=tk.FLAT
        )
        scrollbar = tk.Scrollbar(text_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botões de ação
        btn_frame = tk.Frame(log_frame, bg=self.bg_color)
        btn_frame.pack(pady=(5, 0))
        
        tk.Button(
            btn_frame,
            text="🗑️ Limpar Log",
            command=self.limpar_log,
            bg='#95a5a6',
            fg='white',
            font=('Segoe UI', 8),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            btn_frame,
            text="📋 Copiar Log",
            command=self.copiar_log,
            bg='#95a5a6',
            fg='white',
            font=('Segoe UI', 8),
            cursor='hand2'
        ).pack(side=tk.LEFT, padx=5)
    
    def limpar_log(self):
        """Limpa a área de log"""
        self.log_text.delete(1.0, tk.END)
        self.log("Log limpo com sucesso!", 'info')
    
    def copiar_log(self):
        """Copia o log para a área de transferência"""
        log_content = self.log_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(log_content)
        self.log("Log copiado para a área de transferência!", 'success')
    
    def log(self, mensagem, tipo='info'):
        """Adiciona mensagem ao log com formatação"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        emojis = {
            'error': '❌',
            'success': '✅',
            'warning': '⚠️',
            'info': 'ℹ️'
        }
        
        prefixos = {
            'error': 'ERRO',
            'success': 'SUCESSO',
            'warning': 'AVISO',
            'info': 'INFO'
        }
        
        emoji = emojis.get(tipo, 'ℹ️')
        prefixo = prefixos.get(tipo, 'INFO')
        
        self.log_text.insert(
            tk.END,
            f"[{timestamp}] {emoji} {prefixo}: {mensagem}\n",
            tipo
        )
        self.log_text.see(tk.END)
        self.root.update()
        
        # Configurar tags de cor
        self.log_text.tag_config('error', foreground='#e74c3c')
        self.log_text.tag_config('success', foreground='#27ae60')
        self.log_text.tag_config('warning', foreground='#f39c12')
        self.log_text.tag_config('info', foreground='#3498db')
    
    def selecionar_entrada(self):
        """Abre diálogo para selecionar arquivo XML"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo XML",
            filetypes=[
                ("Arquivos XML", "*.xml"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if arquivo:
            self.arquivo_entrada.set(arquivo)
            nome_base = os.path.splitext(arquivo)[0]
            self.arquivo_saida.set(f"{nome_base}_convertido.xlsx")
            self.status_var.set(f"📄 {os.path.basename(arquivo)}")
            self.log(f"Arquivo selecionado: {arquivo}", 'success')
    
    def selecionar_saida(self):
        """Abre diálogo para selecionar arquivo de saída"""
        arquivo = filedialog.asksaveasfilename(
            title="Salvar arquivo Excel como",
            defaultextension=".xlsx",
            filetypes=[
                ("Arquivos Excel", "*.xlsx"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if arquivo:
            self.arquivo_saida.set(arquivo)
            self.log(f"Arquivo de saída: {arquivo}", 'info')
    
    def abrir_arquivo_excel(self, arquivo):
        """Tenta abrir o arquivo Excel com o programa padrão"""
        try:
            if os.name == 'nt':
                os.startfile(arquivo)
            else:
                import subprocess
                subprocess.run(['open' if os.name == 'darwin' else 'xdg-open', arquivo])
            self.log(f"Arquivo aberto: {arquivo}", 'success')
        except Exception as e:
            self.log(f"Não foi possível abrir o arquivo: {e}", 'warning')
    
    def converter(self):
        """Inicia o processo de conversão"""
        if not self.arquivo_entrada.get():
            messagebox.showerror("Erro", "Selecione um arquivo XML!")
            return
            
        if not os.path.exists(self.arquivo_entrada.get()):
            messagebox.showerror("Erro", "Arquivo não encontrado!")
            return
            
        if not self.arquivo_saida.get():
            messagebox.showerror("Erro", "Defina o arquivo de saída!")
            return
        
        # Desabilita botão e inicia progresso
        self.btn_converter.config(state='disabled', text='🔄 CONVERTENDO...')
        self.progress.start()
        self.status_var.set("Convertendo...")
        self.limpar_log()
        self.log("🚀 Iniciando processo de conversão...", 'info')
        
        # Executa em thread separada
        thread = threading.Thread(target=self.executar_conversao)
        thread.daemon = True
        thread.start()
    
    def executar_conversao(self):
        """Executa a conversão propriamente dita"""
        try:
            arquivo_entrada = self.arquivo_entrada.get()
            arquivo_saida = self.arquivo_saida.get()
            
            self.log(f"📖 Lendo arquivo: {arquivo_entrada}")
            tree = ET.parse(arquivo_entrada)
            root = tree.getroot()
            
            # Detecta o tipo de arquivo
            elementos = root.findall('.//host')
            if elementos:
                self.log(f"🔍 Detectado: Exportação do Zabbix", 'success')
                self.processar_zabbix(root, arquivo_saida)
            else:
                self.log(f"⚠️ Formato não reconhecido, tentando processamento genérico", 'warning')
                self.processar_generico(root, arquivo_saida)
            
        except Exception as e:
            self.log(f"Erro: {str(e)}", 'error')
            import traceback
            self.log(traceback.format_exc(), 'error')
            messagebox.showerror("Erro", f"Falha na conversão:\n{str(e)}")
            
        finally:
            self.btn_converter.config(state='normal', text='🔄 CONVERTER AGORA')
            self.progress.stop()
            self.status_var.set("✅ Conversão finalizada!")
    
    def processar_zabbix(self, root, arquivo_saida):
        """Processa arquivo de exportação do Zabbix"""
        hosts = root.findall('.//host')
        self.log(f"✅ Encontrados {len(hosts)} hosts", 'success')
        
        dados_hosts = []
        
        for idx, host in enumerate(hosts, 1):
            if idx % 50 == 0:
                self.log(f"🔄 Processando {idx}/{len(hosts)} hosts...")
            
            host_data = {}
            
            # Dados básicos
            host_data['host'] = host.findtext('host', '')
            host_data['name'] = host.findtext('name', '')
            host_data['description'] = host.findtext('description', '')
            
            # Template
            template_elem = host.find('.//template/name')
            host_data['template'] = template_elem.text if template_elem is not None else ''
            
            # Groups
            group_elem = host.find('.//group/name')
            host_data['group'] = group_elem.text if group_elem is not None else ''
            
            # Interface
            interface_elem = host.find('.//interface/ip')
            host_data['ip'] = interface_elem.text if interface_elem is not None else ''
            
            # MAC Address
            if host_data['description']:
                mac_match = re.search(r'([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})', host_data['description'])
                host_data['mac'] = mac_match.group(0) if mac_match else ''
            
            # Tags
            for tag in host.findall('.//tag'):
                tag_name = tag.findtext('tag', '')
                tag_value = tag.findtext('value', '')
                if tag_name:
                    host_data[f'tag_{tag_name}'] = tag_value
            
            dados_hosts.append(host_data)
        
        df = pd.DataFrame(dados_hosts)
        
        # Remove colunas vazias se opção marcada
        if self.excluir_vazios.get():
            colunas_antes = len(df.columns)
            df = df.dropna(axis=1, how='all')
            self.log(f"🧹 Removidas {colunas_antes - len(df.columns)} colunas vazias", 'info')
        
        self.salvar_excel(df, arquivo_saida)
    
    def processar_generico(self, root, arquivo_saida):
        """Processamento genérico para qualquer XML"""
        self.log("🔄 Processando XML genérico...", 'info')
        
        dados = []
        
        for elem in root.iter():
            if len(elem) == 0:  # Elementos folha
                dados.append({
                    'tag': elem.tag,
                    'text': elem.text,
                    'atributos': str(elem.attrib) if elem.attrib else ''
                })
        
        df = pd.DataFrame(dados)
        self.salvar_excel(df, arquivo_saida)
    
    def salvar_excel(self, df, arquivo_saida):
        """Salva o DataFrame em Excel"""
        self.log(f"💾 Salvando arquivo: {arquivo_saida}")
        
        with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
            # Dados completos
            df.to_excel(writer, sheet_name='Dados_Completos', index=False)
            self.log(f"✓ Dados completos: {len(df)} linhas x {len(df.columns)} colunas")
            
            # Aba resumida se solicitado
            if self.gerar_resumo.get():
                colunas_importantes = df.columns[:10].tolist()  # Primeiras 10 colunas
                df_resumido = df[colunas_importantes]
                df_resumido.to_excel(writer, sheet_name='Resumo', index=False)
                self.log(f"✓ Aba resumida gerada com {len(colunas_importantes)} campos")
            
            # Ajusta largura das colunas
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        
        self.log("="*50, 'success')
        self.log("✨ CONVERSÃO CONCLUÍDA COM SUCESSO!", 'success')
        self.log("="*50, 'success')
        self.log(f"📁 Arquivo: {arquivo_saida}")
        self.log(f"📊 Total: {len(df)} registros")
        self.log(f"📋 Colunas: {len(df.columns)}")
        
        # Abrir arquivo se opção marcada
        if self.auto_abrir.get():
            self.abrir_arquivo_excel(arquivo_saida)
        
        messagebox.showinfo(
            "Sucesso!",
            f"✨ Conversão concluída com sucesso!\n\n"
            f"📊 {len(df)} registros processados\n"
            f"📋 {len(df.columns)} colunas geradas\n\n"
            f"📁 Arquivo salvo em:\n{arquivo_saida}"
        )

def main():
    """Função principal"""
    root = tk.Tk()
    app = DataCraft(root)
    root.mainloop()

if __name__ == "__main__":
    main()
