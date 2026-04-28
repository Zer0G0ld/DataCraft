"""
Motor Universal de Conversão
Gerencia todo o fluxo de conversão entre formatos
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Any, Optional, Callable
from datetime import datetime
import logging

from .reader_factory import ReaderFactory
from .writer_factory import WriterFactory
from transformers.flatten import DataFlattener
from transformers.mapper import ColumnMapper
from transformers.aggregator import DataAggregator

class ConversionEngine:
    """Motor principal que orquestra todo o processo de conversão"""
    
    def __init__(self):
        self.reader_factory = ReaderFactory()
        self.writer_factory = WriterFactory()
        self.flattener = DataFlattener()
        self.mapper = ColumnMapper()
        self.aggregator = DataAggregator()
        
        self.last_dataframe = None
        self.last_metadata = {}
        self.logger = self._setup_logger()
        self.progress_callback = None
    
    def _setup_logger(self):
        """Configura logger interno"""
        logger = logging.getLogger('DataCraft')
        logger.setLevel(logging.INFO)
        return logger
    
    def set_progress_callback(self, callback: Callable):
        """Define callback para progresso"""
        self.progress_callback = callback
    
    def _update_progress(self, message: str, percentage: int = None):
        """Atualiza progresso via callback"""
        if self.progress_callback:
            self.progress_callback(message, percentage)
        self.logger.info(message)
    
    def convert(self, 
                input_path: str, 
                output_path: str, 
                input_format: Optional[str] = None,
                output_format: Optional[str] = None,
                options: Dict[str, Any] = None) -> Dict[str, Any]:
        """
        Converte qualquer formato para qualquer formato
        
        Args:
            input_path: Caminho do arquivo de entrada
            output_path: Caminho do arquivo de saída  
            input_format: Formato de entrada (auto-detect se None)
            output_format: Formato de saída (detect da extensão se None)
            options: Opções de conversão:
                - flatten (bool): Achatar estruturas aninhadas
                - column_mapping (dict): Mapeamento de colunas
                - filters (dict): Filtros para os dados
                - aggregations (list): Agregações a aplicar
                - reader_options (dict): Opções específicas do leitor
                - writer_options (dict): Opções específicas do escritor
        """
        options = options or {}
        start_time = datetime.now()
        
        try:
            # 1. Detecta formatos
            self._update_progress("🔍 Detectando formatos...", 5)
            if not input_format:
                input_format = self._detect_format(input_path)
            
            if not output_format:
                output_format = self._detect_format_from_extension(output_path)
            
            self._update_progress(f"📖 Formato entrada: {input_format.upper()}", 10)
            self._update_progress(f"💾 Formato saída: {output_format.upper()}", 10)
            
            # 2. Lê o arquivo
            self._update_progress(f"📂 Lendo arquivo {input_format.upper()}...", 15)
            reader = self.reader_factory.get_reader(input_format)
            df = reader.read(input_path, **options.get('reader_options', {}))
            
            self._update_progress(f"✅ Leitura concluída: {len(df)} registros, {len(df.columns)} colunas", 30)
            
            # 3. Aplica transformações
            if options.get('flatten', True):
                self._update_progress("🔄 Achatando estruturas aninhadas...", 40)
                df = self.flattener.flatten_dataframe(df)
                self._update_progress(f"   Estrutura achatada: {len(df.columns)} colunas", 45)
            
            if options.get('column_mapping'):
                self._update_progress("🗺️ Aplicando mapeamento de colunas...", 50)
                df = self.mapper.apply_mapping(df, options['column_mapping'])
            
            if options.get('filters'):
                self._update_progress("🔍 Aplicando filtros...", 60)
                df = self._apply_filters(df, options['filters'])
            
            if options.get('aggregations'):
                self._update_progress("📊 Aplicando agregações...", 70)
                df = self.aggregator.aggregate(df, options['aggregations'])
            
            # 4. Salva no formato de saída
            self._update_progress(f"💾 Salvando como {output_format.upper()}...", 80)
            writer = self.writer_factory.get_writer(output_format)
            result_path = writer.write(df, output_path, **options.get('writer_options', {}))
            
            # 5. Metadados finais
            end_time = datetime.now()
            self.last_dataframe = df
            self.last_metadata = {
                'success': True,
                'input_format': input_format,
                'output_format': output_format,
                'input_path': input_path,
                'output_path': result_path,
                'rows': len(df),
                'columns': len(df.columns),
                'input_size': Path(input_path).stat().st_size,
                'output_size': Path(result_path).stat().st_size if Path(result_path).exists() else 0,
                'duration_seconds': (end_time - start_time).total_seconds(),
                'timestamp': end_time.isoformat()
            }
            
            self._update_progress(f"✅ Conversão concluída em {self.last_metadata['duration_seconds']:.1f}s", 100)
            
            return self.last_metadata
            
        except Exception as e:
            self._update_progress(f"❌ Erro: {str(e)}", 0)
            self.last_metadata = {
                'success': False,
                'error': str(e),
                'timestamp': datetime.now().isoformat()
            }
            raise
    
    def _detect_format(self, file_path: str) -> str:
        """Detecta formato pela extensão e conteúdo"""
        ext = Path(file_path).suffix.lower()
        
        format_map = {
            '.xml': 'xml', '.json': 'json', '.csv': 'csv',
            '.xlsx': 'excel', '.xls': 'excel', '.yaml': 'yaml',
            '.yml': 'yaml', '.parquet': 'parquet', '.sql': 'sql',
            '.html': 'html', '.htm': 'html', '.md': 'markdown'
        }
        
        if ext in format_map:
            return format_map[ext]
        
        # Fallback para detecção por conteúdo
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                first_char = f.read(1).strip()
                if first_char == '<':
                    return 'xml'
                elif first_char in '{[':
                    return 'json'
        except:
            pass
        
        return 'csv'
    
    def _detect_format_from_extension(self, file_path: str) -> str:
        """Detecta formato pela extensão do arquivo de saída"""
        ext = Path(file_path).suffix.lower()
        format_map = {
            '.xlsx': 'excel', '.xls': 'excel', '.csv': 'csv',
            '.json': 'json', '.html': 'html', '.htm': 'html',
            '.md': 'markdown', '.sql': 'sql', '.parquet': 'parquet',
            '.txt': 'csv'
        }
        return format_map.get(ext, 'excel')
    
    def _apply_filters(self, df: pd.DataFrame, filters: Dict) -> pd.DataFrame:
        """Aplica filtros ao DataFrame"""
        filtered_df = df.copy()
        
        for column, condition in filters.items():
            if column in filtered_df.columns:
                if isinstance(condition, dict):
                    if 'equals' in condition:
                        filtered_df = filtered_df[filtered_df[column] == condition['equals']]
                    elif 'contains' in condition:
                        filtered_df = filtered_df[filtered_df[column].astype(str).str.contains(condition['contains'], case=False)]
                    elif 'greater_than' in condition:
                        filtered_df = filtered_df[filtered_df[column] > condition['greater_than']]
                    elif 'less_than' in condition:
                        filtered_df = filtered_df[filtered_df[column] < condition['less_than']]
                else:
                    filtered_df = filtered_df[filtered_df[column] == condition]
        
        return filtered_df
    
    def get_supported_inputs(self) -> list:
        """Retorna lista de formatos de entrada suportados"""
        return self.reader_factory.get_supported_readers()
    
    def get_supported_outputs(self) -> list:
        """Retorna lista de formatos de saída suportados"""
        return self.writer_factory.get_supported_writers()
    
    def preview(self, input_path: str, input_format: Optional[str] = None, rows: int = 10) -> pd.DataFrame:
        """Pré-visualiza os primeiros registros"""
        if not input_format:
            input_format = self._detect_format(input_path)
        
        reader = self.reader_factory.get_reader(input_format)
        df = reader.read(input_path)
        
        return df.head(rows)