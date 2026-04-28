"""
Leitor de arquivos YAML
"""

import pandas as pd
import yaml
from .base_reader import BaseReader

class YAMLReader(BaseReader):
    """Leitor para arquivos YAML"""
    
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        
        # Converte para DataFrame
        if isinstance(data, list):
            return pd.DataFrame(data)
        elif isinstance(data, dict):
            # Procura pela primeira lista
            for key, value in data.items():
                if isinstance(value, list) and value:
                    return pd.DataFrame(value)
            return pd.DataFrame([data])
        else:
            return pd.DataFrame([{'valor': data}])
    
    def get_supported_extensions(self) -> list:
        return ['.yaml', '.yml']