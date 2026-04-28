"""
Transformador para mapeamento de colunas
"""

import pandas as pd
from typing import Dict, List, Optional

class ColumnMapper:
    """Gerencia mapeamento e renomeação de colunas"""
    
    def apply_mapping(self, df: pd.DataFrame, mapping: Dict[str, str]) -> pd.DataFrame:
        """
        Aplica mapeamento de colunas
        
        Args:
            df: DataFrame de entrada
            mapping: Dicionário com {coluna_original: novo_nome}
        
        Returns:
            DataFrame com colunas renomeadas
        """
        df_mapped = df.copy()
        
        # Renomeia colunas existentes
        existing_mapping = {k: v for k, v in mapping.items() if k in df.columns}
        df_mapped = df_mapped.rename(columns=existing_mapping)
        
        return df_mapped
    
    def create_mapping_from_lists(self, source_columns: List[str], 
                                   target_columns: List[str]) -> Dict[str, str]:
        """
        Cria mapeamento a partir de duas listas de colunas
        """
        if len(source_columns) != len(target_columns):
            raise ValueError("Listas devem ter o mesmo tamanho")
        
        return dict(zip(source_columns, target_columns))
    
    def suggest_mapping(self, df: pd.DataFrame, 
                       common_patterns: Optional[List[str]] = None) -> Dict[str, str]:
        """
        Sugere mapeamento baseado em padrões comuns de nomes
        """
        if common_patterns is None:
            common_patterns = ['id', 'name', 'description', 'created_at', 'updated_at']
        
        suggestions = {}
        for col in df.columns:
            col_lower = col.lower()
            for pattern in common_patterns:
                if pattern in col_lower:
                    suggestions[col] = pattern
                    break
        
        return suggestions