"""
Classe Base para todos os escritores
"""

from abc import ABC, abstractmethod
import pandas as pd
from typing import Dict, Any

class BaseWriter(ABC):
    """Classe base abstrata para todos os escritores"""
    
    @abstractmethod
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        """
        Escreve o DataFrame em um arquivo
        
        Args:
            df: DataFrame a ser escrito
            output_path: Caminho do arquivo de saída
            **kwargs: Opções específicas do escritor
        
        Returns:
            Caminho do arquivo escrito
        """
        pass
    
    @abstractmethod
    def get_supported_formats(self) -> list:
        """Retorna lista de formatos suportados"""
        pass