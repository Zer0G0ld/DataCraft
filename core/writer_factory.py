"""
Fábrica de Escritores - Gerencia todos os escritores disponíveis
"""

from typing import Dict, Type
from writers.base_writer import BaseWriter

class WriterFactory:
    """Fábrica para criar escritores de diferentes formatos"""
    
    def __init__(self):
        self._writers: Dict[str, Type[BaseWriter]] = {}
        self._register_default_writers()
    
    def _register_default_writers(self):
        """Registra todos os escritores disponíveis"""
        from writers.excel_writer import ExcelWriter
        from writers.json_writer import JSONWriter
        from writers.csv_writer import CSVWriter
        from writers.html_writer import HTMLWriter
        from writers.markdown_writer import MarkdownWriter
        from writers.sql_writer import SQLWriter
        from writers.parquet_writer import ParquetWriter
        
        self.register_writer('excel', ExcelWriter)
        self.register_writer('json', JSONWriter)
        self.register_writer('csv', CSVWriter)
        self.register_writer('html', HTMLWriter)
        self.register_writer('markdown', MarkdownWriter)
        self.register_writer('sql', SQLWriter)
        self.register_writer('parquet', ParquetWriter)
    
    def register_writer(self, format_name: str, writer_class: Type[BaseWriter]):
        """Registra um novo escritor"""
        self._writers[format_name.lower()] = writer_class
    
    def get_writer(self, format_name: str) -> BaseWriter:
        """Retorna uma instância do escritor para o formato"""
        format_name = format_name.lower()
        
        # Mapeia sinônimos
        synonyms = {
            'xlsx': 'excel',
            'xls': 'excel',
            'pq': 'parquet',
            'md': 'markdown',
            'htm': 'html'
        }
        
        format_name = synonyms.get(format_name, format_name)
        
        if format_name not in self._writers:
            raise ValueError(f"Formato não suportado: {format_name}. "
                           f"Suportados: {list(self._writers.keys())}")
        
        writer_class = self._writers[format_name]
        return writer_class()
    
    def get_supported_writers(self) -> list:
        """Retorna lista de formatos suportados"""
        return list(self._writers.keys())
    
    def get_writer_info(self, format_name: str) -> dict:
        """Retorna informações sobre um escritor específico"""
        writer = self.get_writer(format_name)
        return {
            'name': format_name,
            'formats': writer.get_supported_formats(),
            'class': writer.__class__.__name__
        }