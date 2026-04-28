"""
Leitor de arquivos CSV
"""

import pandas as pd
from .base_reader import BaseReader

class CSVReader(BaseReader):
    """Leitor para arquivos CSV"""
    
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        """
        Lê arquivo CSV com detecção automática de delimitador
        """
        # Tenta detectar o delimitador
        delimiters = [',', ';', '\t', '|']
        encoding = kwargs.get('encoding', 'utf-8')
        
        for delimiter in delimiters:
            try:
                df = pd.read_csv(file_path, delimiter=delimiter, encoding=encoding, nrows=5)
                if len(df.columns) > 1:
                    # Leitura completa com o delimitador encontrado
                    return pd.read_csv(file_path, delimiter=delimiter, encoding=encoding, **kwargs)
            except:
                continue
        
        # Fallback: tenta com detecção automática do pandas
        return pd.read_csv(file_path, encoding=encoding, **kwargs)
    
    def get_supported_extensions(self) -> list:
        return ['.csv', '.txt']