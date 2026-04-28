"""
Transformador para achatamento de estruturas aninhadas
"""

import pandas as pd
from typing import Dict, Any, List

class DataFlattener:
    """Achata estruturas aninhadas (listas e dicionários) em DataFrames"""
    
    def flatten_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Achata um DataFrame inteiro, expandindo estruturas aninhadas
        
        Args:
            df: DataFrame de entrada
        
        Returns:
            DataFrame achatado
        """
        # Se já está achatado, retorna
        if not any(df.dtypes == 'object'):
            return df
        
        # Aplica flattening em cada linha
        flattened_data = []
        for idx, row in df.iterrows():
            flattened_row = self._flatten_row(row.to_dict())
            flattened_data.append(flattened_row)
        
        result_df = pd.DataFrame(flattened_data)
        
        # Remove colunas completamente vazias
        result_df = result_df.dropna(axis=1, how='all')
        
        return result_df
    
    def _flatten_row(self, row: Dict[str, Any], parent_key: str = '') -> Dict[str, Any]:
        """Achata uma linha/objeto recursivamente"""
        items = {}
        
        for key, value in row.items():
            new_key = f"{parent_key}.{key}" if parent_key else key
            new_key = new_key.replace(' ', '_').lower()
            
            if isinstance(value, dict):
                items.update(self._flatten_row(value, new_key))
            elif isinstance(value, list):
                if value and isinstance(value[0], dict):
                    # Lista de dicionários
                    for i, item in enumerate(value):
                        items.update(self._flatten_row(item, f"{new_key}_{i+1}"))
                else:
                    # Lista de valores simples
                    items[new_key] = ', '.join(str(x) for x in value if x is not None) if value else ''
            elif isinstance(value, (int, float, str, bool)) or value is None:
                items[new_key] = value if value is not None and pd.notna(value) else ''
            else:
                items[new_key] = str(value) if value else ''
        
        return items
    
    def flatten_single_object(self, obj: Dict[str, Any], prefix: str = '') -> Dict[str, Any]:
        """Achata um único objeto/dicionário"""
        return self._flatten_row(obj, prefix)