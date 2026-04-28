"""
Escritor para arquivos CSV
"""

import pandas as pd
from .base_writer import BaseWriter

class CSVWriter(BaseWriter):
    """Escritor para formato CSV"""
    
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        delimiter = kwargs.get('delimiter', ';')
        encoding = kwargs.get('encoding', 'utf-8-sig')
        index = kwargs.get('index', False)
        
        df.to_csv(output_path, sep=delimiter, encoding=encoding, index=index)
        
        return output_path
    
    def get_supported_formats(self) -> list:
        return ['csv']