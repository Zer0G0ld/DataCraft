"""
Classe Base para todos os leitores
"""

from abc import ABC, abstractmethod
import pandas as pd
from typing import Dict, Any, Optional

class BaseReader(ABC):
    """Classe base abstrata para todos os leitores de arquivos"""
    
    @abstractmethod
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        """
        Lê o arquivo e retorna um DataFrame
        
        Args:
            file_path: Caminho do arquivo
            **kwargs: Opções específicas do leitor
        
        Returns:
            DataFrame com os dados lidos
        """
        pass
    
    @abstractmethod
    def get_supported_extensions(self) -> list:
        """Retorna lista de extensões suportadas"""
        pass
    
    def validate(self, file_path: str) -> bool:
        """Valida se o arquivo pode ser lido"""
        try:
            test_df = self.read(file_path, nrows=1)
            return True
        except:
            return False
    
    def get_info(self, file_path: str) -> Dict[str, Any]:
        """Retorna informações sobre o arquivo sem carregar tudo"""
        df = self.read(file_path)
        return {
            'rows': len(df),
            'columns': len(df.columns),
            'column_names': df.columns.tolist(),
            'memory_usage': df.memory_usage(deep=True).sum()
        }