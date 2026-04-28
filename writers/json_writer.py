"""
Escritor para arquivos JSON
"""

import pandas as pd
import json
from .base_writer import BaseWriter

class JSONWriter(BaseWriter):
    """Escritor para formato JSON"""
    
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        orient = kwargs.get('orient', 'records')  # records, index, columns, values
        indent = kwargs.get('indent', 2)
        
        # Converte DataFrame para dicionário
        data = df.to_dict(orient=orient)
        
        # Salva como JSON
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=indent, ensure_ascii=False)
        
        return output_path
    
    def get_supported_formats(self) -> list:
        return ['json']