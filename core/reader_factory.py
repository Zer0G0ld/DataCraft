"""
Fábrica de Leitores - Gerencia todos os leitores disponíveis
"""

from typing import Dict, Type
from readers.base_reader import BaseReader

class ReaderFactory:
    """Fábrica para criar leitores de diferentes formatos"""
    
    def __init__(self):
        self._readers: Dict[str, Type[BaseReader]] = {}
        self._register_default_readers()
    
    def _register_default_readers(self):
        """Registra todos os leitores disponíveis"""
        from readers.xml_reader import XMLReader
        from readers.json_reader import JSONReader
        from readers.csv_reader import CSVReader
        from readers.yaml_reader import YAMLReader
        from readers.excel_reader import ExcelReader
        from readers.sql_reader import SQLReader
        from readers.parquet_reader import ParquetReader
        
        self.register_reader('xml', XMLReader)
        self.register_reader('json', JSONReader)
        self.register_reader('csv', CSVReader)
        self.register_reader('yaml', YAMLReader)
        self.register_reader('excel', ExcelReader)
        self.register_reader('sql', SQLReader)
        self.register_reader('parquet', ParquetReader)
    
    def register_reader(self, format_name: str, reader_class: Type[BaseReader]):
        """Registra um novo leitor"""
        self._readers[format_name.lower()] = reader_class
    
    def get_reader(self, format_name: str) -> BaseReader:
        """Retorna uma instância do leitor para o formato"""
        format_name = format_name.lower()
        
        # Mapeia sinônimos
        synonyms = {
            'xlsx': 'excel',
            'xls': 'excel',
            'yml': 'yaml',
            'pq': 'parquet',
            'sqlite': 'sql',
            'db': 'sql'
        }
        
        format_name = synonyms.get(format_name, format_name)
        
        if format_name not in self._readers:
            raise ValueError(f"Formato não suportado: {format_name}. "
                           f"Suportados: {list(self._readers.keys())}")
        
        reader_class = self._readers[format_name]
        return reader_class()
    
    def get_supported_readers(self) -> list:
        """Retorna lista de formatos suportados"""
        return list(self._readers.keys())
    
    def get_reader_info(self, format_name: str) -> dict:
        """Retorna informações sobre um leitor específico"""
        reader = self.get_reader(format_name)
        return {
            'name': format_name,
            'extensions': reader.get_supported_extensions(),
            'class': reader.__class__.__name__
        }