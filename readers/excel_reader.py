"""
Leitor de arquivos Excel
"""

import pandas as pd
from .base_reader import BaseReader

class ExcelReader(BaseReader):
    """Leitor para arquivos Excel (.xlsx, .xls)"""
    
    def read(self, file_path: str, **kwargs) -> pd.DataFrame:
        """
        Lê arquivo Excel
        
        Args:
            file_path: Caminho do arquivo
            **kwargs:
                - sheet_name: Nome da aba (padrão: primeira aba)
                - header: Linha para cabeçalho (padrão: 0)
                - skiprows: Linhas pular
                - nrows: Número de linhas para ler
        """
        sheet_name = kwargs.get('sheet_name', 0)  # 0 = primeira aba
        header = kwargs.get('header', 0)
        skiprows = kwargs.get('skiprows', None)
        nrows = kwargs.get('nrows', None)
        
        # Detecta engine baseado na extensão
        if file_path.endswith('.xls'):
            engine = 'xlrd'
        else:
            engine = 'openpyxl'
        
        try:
            df = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                header=header,
                skiprows=skiprows,
                nrows=nrows,
                engine=engine
            )
            return df
        except Exception as e:
            # Fallback: tenta sem especificar engine
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)
            return df
    
    def get_sheets(self, file_path: str) -> list:
        """Retorna lista de abas do Excel"""
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    
    def read_all_sheets(self, file_path: str, **kwargs) -> dict:
        """Lê todas as abas e retorna dicionário {sheet_name: dataframe}"""
        excel_file = pd.ExcelFile(file_path)
        all_sheets = {}
        
        for sheet in excel_file.sheet_names:
            all_sheets[sheet] = pd.read_excel(file_path, sheet_name=sheet, **kwargs)
        
        return all_sheets
    
    def get_supported_extensions(self) -> list:
        return ['.xlsx', '.xls', '.xlsm']