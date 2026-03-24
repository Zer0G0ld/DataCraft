"""
DataCraft - Conversor Inteligente de XML/JSON para Excel
Desenvolvido por Zer0G0ld 🏆
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import xml.etree.ElementTree as ET
import json
import pandas as pd
import os
import re
import sys
from datetime import datetime

class DataCraft:
    """Aplicativo principal para conversão de XML/JSON para Excel"""
    
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
        self.formato_arquivo = tk.StringVar(value="auto")  # auto, xml, json
        
        # Criar interface
        self.criar_interface()
        
        # Centralizar janela
        self.centralizar_janela()
        
        # Log de inicialização
        self.log("DataCraft iniciado com sucesso!", 'success')
        self.log("Selecione um arquivo XML ou JSON para começar...", 'info')
    
    def configurar_icone(self):
        """Configura o ícone da aplicação"""
        try:
            icon_paths = []
            
            if os.path.exists('voto.ico'):
                icon_paths.append('voto.ico')
            
            if getattr(sys, 'frozen', False):
                application_path = os.path.dirname(sys.executable)
                icon_path = os.path.join(application_path, 'voto.ico')
                if os.path.exists(icon_path):
                    icon_paths.append(icon_path)
            
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
        main_frame = tk.Frame(self.root, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.criar_titulo(main_frame)
        self.criar_frame_arquivos(main_frame)
        self.criar_opcoes(main_frame)
        self.criar_botao_converter(main_frame)
        self.criar_status_progresso(main_frame)
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
            text="Converta arquivos XML ou JSON para Excel de forma rápida e inteligente",
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
            text="Arquivo:",
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
        
        # Formato do arquivo
        formato_frame = tk.Frame(arquivos_frame, bg=self.bg_color)
        formato_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(
            formato_frame,
            text="Formato:",
            width=12,
            anchor='w',
            font=('Segoe UI', 10),
            bg=self.bg_color
        ).pack(side=tk.LEFT)
        
        tk.Radiobutton(
            formato_frame,
            text="Auto-detectar",
            variable=self.formato_arquivo,
            value="auto",
            bg=self.bg_color,
            font=('Segoe UI', 9)
        ).pack(side=tk.LEFT, padx=(5, 5))
        
        tk.Radiobutton(
            formato_frame,
            text="XML",
            variable=self.formato_arquivo,
            value="xml",
            bg=self.bg_color,
            font=('Segoe UI', 9)
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Radiobutton(
            formato_frame,
            text="JSON",
            variable=self.formato_arquivo,
            value="json",
            bg=self.bg_color,
            font=('Segoe UI', 9)
        ).pack(side=tk.LEFT, padx=5)
        
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
        
        self.normalizar_json = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opcoes_frame,
            text="🔄 Expandir estruturas aninhadas",
            variable=self.normalizar_json,
            bg=self.bg_color,
            font=('Segoe UI', 9)
        ).pack(anchor=tk.W)
        
        self.extrair_mac = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opcoes_frame,
            text="🔍 Tentar extrair MAC Address de campos de texto",
            variable=self.extrair_mac,
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
        
        self.log_text.tag_config('error', foreground='#e74c3c')
        self.log_text.tag_config('success', foreground='#27ae60')
        self.log_text.tag_config('warning', foreground='#f39c12')
        self.log_text.tag_config('info', foreground='#3498db')
    
    def extrair_mac_address(self, texto):
        """Extrai MAC Address de um texto se a opção estiver ativada"""
        if not self.extrair_mac.get() or not texto:
            return ''
        mac_match = re.search(r'([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})', str(texto), re.IGNORECASE)
        return mac_match.group(0) if mac_match else ''
    
    def detectar_formato(self, arquivo):
        """Detecta o formato do arquivo baseado na extensão e conteúdo"""
        if self.formato_arquivo.get() != "auto":
            return self.formato_arquivo.get()
        
        extensao = os.path.splitext(arquivo)[1].lower()
        if extensao == '.xml':
            return 'xml'
        elif extensao == '.json':
            return 'json'
        elif extensao == '.yaml' or extensao == '.yml':
            return 'yaml'
        else:
            try:
                with open(arquivo, 'r', encoding='utf-8') as f:
                    primeiro_caractere = f.read(1).strip()
                    if primeiro_caractere == '<':
                        return 'xml'
                    elif primeiro_caractere == '{' or primeiro_caractere == '[':
                        return 'json'
            except:
                pass
            return 'xml'
    
    def selecionar_entrada(self):
        """Abre diálogo para selecionar arquivo"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo XML ou JSON",
            filetypes=[
                ("Arquivos suportados", "*.xml;*.json"),
                ("Arquivos XML", "*.xml"),
                ("Arquivos JSON", "*.json"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if arquivo:
            self.arquivo_entrada.set(arquivo)
            nome_base = os.path.splitext(arquivo)[0]
            # Remove espaços e caracteres especiais do nome base
            nome_base_limpo = re.sub(r'[<>:"/\\|?*()\[\]\s]+', '_', nome_base)
            nome_base_limpo = re.sub(r'_+', '_', nome_base_limpo)
            self.arquivo_saida.set(f"{nome_base_limpo}_convertido.xlsx")
            self.status_var.set(f"📄 {os.path.basename(arquivo)}")
            formato = self.detectar_formato(arquivo)
            self.log(f"Arquivo selecionado: {arquivo} (formato: {formato.upper()})", 'success')
    
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
            messagebox.showerror("Erro", "Selecione um arquivo!")
            return
            
        if not os.path.exists(self.arquivo_entrada.get()):
            messagebox.showerror("Erro", "Arquivo não encontrado!")
            return
            
        if not self.arquivo_saida.get():
            messagebox.showerror("Erro", "Defina o arquivo de saída!")
            return
        
        self.btn_converter.config(state='disabled', text='🔄 CONVERTENDO...')
        self.progress.start()
        self.status_var.set("Convertendo...")
        self.limpar_log()
        self.log("🚀 Iniciando processo de conversão...", 'info')
        
        thread = threading.Thread(target=self.executar_conversao)
        thread.daemon = True
        thread.start()
    
    def executar_conversao(self):
        """Executa a conversão propriamente dita"""
        try:
            arquivo_entrada = self.arquivo_entrada.get()
            arquivo_saida = self.arquivo_saida.get()
            
            formato = self.detectar_formato(arquivo_entrada)
            self.log(f"📖 Lendo arquivo: {arquivo_entrada} (formato: {formato.upper()})", 'info')
            
            if formato == 'xml':
                self.processar_xml(arquivo_entrada, arquivo_saida)
            else:
                self.processar_json(arquivo_entrada, arquivo_saida)
            
        except Exception as e:
            self.log(f"Erro: {str(e)}", 'error')
            import traceback
            self.log(traceback.format_exc(), 'error')
            messagebox.showerror("Erro", f"Falha na conversão:\n{str(e)}")
            
        finally:
            self.btn_converter.config(state='normal', text='🔄 CONVERTER AGORA')
            self.progress.stop()
            self.status_var.set("✅ Conversão finalizada!")
    
    def processar_xml(self, arquivo_entrada, arquivo_saida):
        """Processa arquivo XML - achata todas as estruturas aninhadas em uma única planilha"""
        self.log("🔍 Processando arquivo XML...", 'info')
        tree = ET.parse(arquivo_entrada)
        root = tree.getroot()
        
        # Função para converter elemento XML em dicionário
        def element_to_dict(element, prefixo=''):
            """Converte um elemento XML em dicionário, tratando elementos repetidos como listas"""
            resultado = {}
            
            # Conta ocorrências de cada tag filho
            contagem_tags = {}
            for child in element:
                tag = child.tag
                contagem_tags[tag] = contagem_tags.get(tag, 0) + 1
            
            # Processa cada filho
            for child in element:
                tag = child.tag
                novo_prefixo = f"{prefixo}{tag}" if prefixo else tag
                
                # Se tem texto e não tem filhos, é um valor simples
                if len(child) == 0:
                    valor = child.text.strip() if child.text else ''
                    if contagem_tags[tag] > 1:
                        # Se a tag se repete, usa lista
                        if novo_prefixo not in resultado:
                            resultado[novo_prefixo] = []
                        resultado[novo_prefixo].append(valor)
                    else:
                        resultado[novo_prefixo] = valor
                else:
                    # Tem filhos, processa recursivamente
                    if contagem_tags[tag] > 1:
                        # Tag repetida - cria lista de objetos
                        if novo_prefixo not in resultado:
                            resultado[novo_prefixo] = []
                        resultado[novo_prefixo].append(element_to_dict(child))
                    else:
                        # Tag única - objeto simples
                        resultado.update(element_to_dict(child, f"{novo_prefixo}."))
            
            return resultado
        
        # Converte o XML inteiro para dicionário
        self.log("🔄 Convertendo XML para dicionário...", 'info')
        dados = element_to_dict(root)
        
        # Encontra todas as listas no dicionário e pega a maior
        listas_encontradas = []
        
        def encontrar_todas_listas(obj, caminho):
            if isinstance(obj, list) and len(obj) > 0:
                listas_encontradas.append({
                    'caminho': caminho.copy(),
                    'tamanho': len(obj),
                    'dados': obj
                })
            elif isinstance(obj, dict):
                for key, value in obj.items():
                    caminho.append(key)
                    encontrar_todas_listas(value, caminho)
                    caminho.pop()
        
        encontrar_todas_listas(dados, [])
        
        if not listas_encontradas:
            self.log("⚠️ Nenhuma lista encontrada no XML", 'warning')
            # Tenta processar como objeto único
            if isinstance(dados, dict):
                dados = [dados]
            else:
                dados = [{'valor': str(dados)}]
            df = pd.DataFrame(dados)
            self.salvar_excel(df, arquivo_saida)
            return
        
        # Ordena por tamanho e pega a maior lista
        listas_encontradas.sort(key=lambda x: x['tamanho'], reverse=True)
        lista_principal = listas_encontradas[0]
        
        caminho_str = ' → '.join(map(str, lista_principal['caminho']))
        self.log(f"📋 Lista principal: '{caminho_str}' com {lista_principal['tamanho']} itens", 'success')
        
        dados_processar = lista_principal['dados']
        
        # Função para achatamento recursivo (mesma do JSON)
        def achatar_objeto(obj, prefixo=''):
            """Achata um objeto/dicionário em um dicionário simples"""
            resultado = {}
            
            if isinstance(obj, dict):
                for chave, valor in obj.items():
                    novo_prefixo = f"{prefixo}{chave}" if prefixo else chave
                    
                    if isinstance(valor, dict):
                        resultado.update(achatar_objeto(valor, f"{novo_prefixo}."))
                    elif isinstance(valor, list):
                        if len(valor) == 0:
                            resultado[novo_prefixo] = ''
                        elif len(valor) == 1 and isinstance(valor[0], (str, int, float, bool, type(None))):
                            resultado[novo_prefixo] = valor[0] if valor[0] is not None else ''
                        elif all(isinstance(item, (str, int, float, bool, type(None))) for item in valor):
                            valores = [str(v) for v in valor if v is not None]
                            resultado[novo_prefixo] = ', '.join(valores) if valores else ''
                        else:
                            for i, item in enumerate(valor, 1):
                                if isinstance(item, dict):
                                    for sub_chave, sub_valor in achatar_objeto(item, f"{novo_prefixo}_{i}.").items():
                                        resultado[sub_chave] = sub_valor
                                elif item is not None:
                                    resultado[f"{novo_prefixo}_{i}"] = str(item)
                    else:
                        resultado[novo_prefixo] = valor if valor is not None else ''
            else:
                resultado[prefixo or 'valor'] = str(obj) if obj is not None else ''
            
            return resultado
        
        # Processa cada item da lista principal
        self.log("🔄 Achatando estruturas aninhadas...", 'info')
        dados_achatados = []
        
        for idx, item in enumerate(dados_processar):
            if (idx + 1) % 10 == 0 or idx == 0:
                self.log(f"   Processando item {idx + 1}/{len(dados_processar)}...")
            
            if isinstance(item, dict):
                item_achatado = achatar_objeto(item)
            else:
                item_achatado = {'valor': str(item) if item is not None else ''}
            
            dados_achatados.append(item_achatado)
        
        # Cria DataFrame
        df = pd.DataFrame(dados_achatados)
        
        self.log(f"✅ Dados carregados: {len(df)} registros, {len(df.columns)} colunas", 'success')
        
        # Extrai MAC Address se opção ativada
        if self.extrair_mac.get():
            self.log("🔍 Procurando MAC Address nos campos de texto...", 'info')
            mac_encontrados = 0
            for col in df.columns:
                if df[col].dtype == 'object':
                    mac_col = df[col].apply(lambda x: self.extrair_mac_address(str(x)) if x else '')
                    if mac_col.any():
                        df[col + '_MAC'] = mac_col
                        mac_encontrados += mac_col.astype(bool).sum()
            if mac_encontrados > 0:
                self.log(f"🔍 Encontrados {mac_encontrados} MAC Addresses", 'success')
        
        # Remove colunas vazias
        if self.excluir_vazios.get():
            colunas_antes = len(df.columns)
            df = df.dropna(axis=1, how='all')
            removidas = colunas_antes - len(df.columns)
            if removidas > 0:
                self.log(f"🧹 Removidas {removidas} colunas vazias", 'info')
        
        # Reordena colunas para deixar as principais primeiro
        colunas_principais = ['host', 'name', 'ip', 'description', 'mac']
        colunas_existentes = [col for col in colunas_principais if col in df.columns]
        outras_colunas = [col for col in df.columns if col not in colunas_existentes]
        if colunas_existentes:
            df = df[colunas_existentes + outras_colunas]
        
        # Mostra estatísticas no log
        self.log(f"\n📊 Estatísticas dos dados:")
        self.log(f"   - Total de registros: {len(df)}")
        if 'host' in df.columns:
            self.log(f"   - Hosts únicos: {df['host'].nunique()}")
        
        self.salvar_excel(df, arquivo_saida)

    def gerar_nome_arquivo_sem_caracteres_especiais(self, caminho):
        """Remove caracteres especiais do nome do arquivo"""
        diretorio = os.path.dirname(caminho)
        nome_arquivo = os.path.basename(caminho)
        
        # Remove caracteres problemáticos para Windows
        nome_limpo = re.sub(r'[<>:"/\\|?*()\[\]\s]+', '_', nome_arquivo)
        # Remove underscores duplicados
        nome_limpo = re.sub(r'_+', '_', nome_limpo)
        
        return os.path.join(diretorio, nome_limpo)
    
    def processar_json(self, arquivo_entrada, arquivo_saida):
        """Processa arquivo JSON - achata todas as estruturas aninhadas em uma única planilha"""
        self.log("🔍 Processando arquivo JSON...", 'info')
        
        with open(arquivo_entrada, 'r', encoding='utf-8') as f:
            dados = json.load(f)
        
        # Encontra todas as listas no JSON e pega a maior
        listas_encontradas = []
        
        def encontrar_todas_listas(obj, caminho):
            if isinstance(obj, list) and len(obj) > 0:
                listas_encontradas.append({
                    'caminho': caminho.copy(),
                    'tamanho': len(obj),
                    'dados': obj
                })
            elif isinstance(obj, dict):
                for key, value in obj.items():
                    caminho.append(key)
                    encontrar_todas_listas(value, caminho)
                    caminho.pop()
        
        encontrar_todas_listas(dados, [])
        
        if not listas_encontradas:
            self.log("⚠️ Nenhuma lista encontrada no JSON", 'warning')
            self.salvar_excel(pd.DataFrame(), arquivo_saida)
            return
        
        # Ordena por tamanho e pega a maior lista
        listas_encontradas.sort(key=lambda x: x['tamanho'], reverse=True)
        lista_principal = listas_encontradas[0]
        
        caminho_str = ' → '.join(map(str, lista_principal['caminho']))
        self.log(f"📋 Lista principal: '{caminho_str}' com {lista_principal['tamanho']} itens", 'success')
        
        dados_processar = lista_principal['dados']
        
        # Função para achatamento recursivo
        def achatar_objeto(obj, prefixo=''):
            """Achata um objeto/dicionário em um dicionário simples"""
            resultado = {}
            
            if isinstance(obj, dict):
                for chave, valor in obj.items():
                    novo_prefixo = f"{prefixo}{chave}" if prefixo else chave
                    
                    if isinstance(valor, dict):
                        # Dicionário aninhado - expande com ponto
                        resultado.update(achatar_objeto(valor, f"{novo_prefixo}."))
                    elif isinstance(valor, list):
                        # Lista - processa cada item
                        if len(valor) == 0:
                            resultado[novo_prefixo] = ''
                        elif len(valor) == 1 and isinstance(valor[0], (str, int, float, bool, type(None))):
                            # Lista com um único valor simples
                            resultado[novo_prefixo] = valor[0] if valor[0] is not None else ''
                        elif all(isinstance(item, (str, int, float, bool, type(None))) for item in valor):
                            # Lista de valores simples - junta com vírgula
                            valores = [str(v) for v in valor if v is not None]
                            resultado[novo_prefixo] = ', '.join(valores) if valores else ''
                        else:
                            # Lista de objetos - expande cada um
                            for i, item in enumerate(valor, 1):
                                if isinstance(item, dict):
                                    for sub_chave, sub_valor in achatar_objeto(item, f"{novo_prefixo}_{i}.").items():
                                        resultado[sub_chave] = sub_valor
                                elif item is not None:
                                    resultado[f"{novo_prefixo}_{i}"] = str(item)
                    else:
                        # Valor simples
                        resultado[novo_prefixo] = valor if valor is not None else ''
            else:
                resultado[prefixo or 'valor'] = str(obj) if obj is not None else ''
            
            return resultado
        
        # Processa cada item da lista principal
        self.log("🔄 Achatando estruturas aninhadas...", 'info')
        dados_achatados = []
        
        for idx, item in enumerate(dados_processar):
            if (idx + 1) % 10 == 0 or idx == 0:
                self.log(f"   Processando item {idx + 1}/{len(dados_processar)}...")
            
            if isinstance(item, dict):
                item_achatado = achatar_objeto(item)
            else:
                item_achatado = {'valor': str(item) if item is not None else ''}
            
            dados_achatados.append(item_achatado)
        
        # Cria DataFrame
        df = pd.DataFrame(dados_achatados)
        
        self.log(f"✅ Dados carregados: {len(df)} registros, {len(df.columns)} colunas", 'success')
        
        # Extrai MAC Address se opção ativada
        if self.extrair_mac.get():
            self.log("🔍 Procurando MAC Address nos campos de texto...", 'info')
            mac_encontrados = 0
            for col in df.columns:
                if df[col].dtype == 'object':
                    mac_col = df[col].apply(lambda x: self.extrair_mac_address(str(x)) if x else '')
                    if mac_col.any():
                        df[col + '_MAC'] = mac_col
                        mac_encontrados += mac_col.astype(bool).sum()
            if mac_encontrados > 0:
                self.log(f"🔍 Encontrados {mac_encontrados} MAC Addresses", 'success')
        
        # Remove colunas vazias
        if self.excluir_vazios.get():
            colunas_antes = len(df.columns)
            df = df.dropna(axis=1, how='all')
            removidas = colunas_antes - len(df.columns)
            if removidas > 0:
                self.log(f"🧹 Removidas {removidas} colunas vazias", 'info')
        
        # Reordena colunas para deixar as principais primeiro
        colunas_principais = ['host', 'name', 'ip', 'description']
        colunas_existentes = [col for col in colunas_principais if col in df.columns]
        outras_colunas = [col for col in df.columns if col not in colunas_existentes]
        df = df[colunas_existentes + outras_colunas]
        
        # Mostra estatísticas no log
        self.log(f"\n📊 Estatísticas dos dados:")
        self.log(f"   - Total de registros: {len(df)}")
        
        # Mostra algumas colunas importantes se existirem
        colunas_interessantes = ['tags.Criticidade', 'tags.Fabricante', 'tags.Pavimento', 'tags.Setor', 'tags.Tipo']
        colunas_comuns = [col for col in colunas_interessantes if col in df.columns]
        if colunas_comuns:
            self.log(f"\n📌 Distribuições:")
            for col in colunas_comuns[:3]:  # Mostra até 3 distribuições
                valores = df[col].value_counts()
                if len(valores) > 0:
                    campo = col.split('.')[-1] if '.' in col else col
                    self.log(f"   {campo}:")
                    for val, count in valores.head(5).items():
                        self.log(f"      - {val}: {count}")
        
        self.salvar_excel(df, arquivo_saida)
    
    def salvar_excel(self, df, arquivo_saida):
        """Salva o DataFrame em Excel com tratamento de erros"""
        if df.empty:
            self.log("⚠️ Nenhum dado para salvar!", 'warning')
            messagebox.showwarning("Aviso", "Nenhum dado foi extraído do arquivo!")
            return
        
        # Limpa o nome do arquivo
        arquivo_saida_original = arquivo_saida
        arquivo_saida = self.gerar_nome_arquivo_sem_caracteres_especiais(arquivo_saida)
        
        if arquivo_saida != arquivo_saida_original:
            self.log(f"📝 Nome do arquivo ajustado: {os.path.basename(arquivo_saida)}", 'info')
        
        # Verifica se o arquivo já existe e tenta liberar
        if os.path.exists(arquivo_saida):
            try:
                # Tenta abrir o arquivo para escrita (testa se está bloqueado)
                with open(arquivo_saida, 'a'):
                    pass
            except PermissionError:
                # Arquivo bloqueado, cria um nome alternativo
                base, ext = os.path.splitext(arquivo_saida)
                contador = 1
                while os.path.exists(f"{base}_{contador}{ext}"):
                    contador += 1
                arquivo_saida = f"{base}_{contador}{ext}"
                self.log(f"⚠️ Arquivo original estava aberto, salvando como: {os.path.basename(arquivo_saida)}", 'warning')
        
        self.log(f"💾 Salvando arquivo: {arquivo_saida}")
        
        try:
            with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Dados_Completos', index=False)
                self.log(f"✓ Dados completos: {len(df)} linhas x {len(df.columns)} colunas")
                
                if self.gerar_resumo.get() and len(df.columns) > 0:
                    # Pega as primeiras 10 colunas ou colunas importantes
                    colunas_importantes = ['host', 'name', 'ip', 'mac', 'description']
                    colunas_existentes = [col for col in colunas_importantes if col in df.columns]
                    
                    if not colunas_existentes:
                        colunas_existentes = df.columns[:min(10, len(df.columns))].tolist()
                    
                    df_resumido = df[colunas_existentes]
                    df_resumido.to_excel(writer, sheet_name='Resumo', index=False)
                    self.log(f"✓ Aba resumida gerada com {len(colunas_existentes)} campos")
                
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
            
            if self.auto_abrir.get():
                self.abrir_arquivo_excel(arquivo_saida)
            
            messagebox.showinfo(
                "Sucesso!",
                f"✨ Conversão concluída com sucesso!\n\n"
                f"📊 {len(df)} registros processados\n"
                f"📋 {len(df.columns)} colunas geradas\n\n"
                f"📁 Arquivo salvo em:\n{arquivo_saida}"
            )
            
        except PermissionError as e:
            self.log(f"❌ Erro de permissão: O arquivo pode estar aberto em outro programa", 'error')
            self.log(f"   Feche o arquivo se ele estiver aberto no Excel e tente novamente", 'error')
            messagebox.showerror(
                "Erro de Permissão",
                f"Não foi possível salvar o arquivo!\n\n"
                f"O arquivo pode estar aberto no Excel ou outro programa.\n\n"
                f"Feche o arquivo e tente novamente.\n\n"
                f"Arquivo: {arquivo_saida}"
            )
        except Exception as e:
            self.log(f"❌ Erro ao salvar: {str(e)}", 'error')
            messagebox.showerror("Erro", f"Falha ao salvar arquivo:\n{str(e)}")

def main():
    """Função principal"""
    root = tk.Tk()
    app = DataCraft(root)
    root.mainloop()

if __name__ == "__main__":
    main()
