"""
Escritor para arquivos Markdown
"""

import pandas as pd
from .base_writer import BaseWriter

class MarkdownWriter(BaseWriter):
    """Escritor para formato Markdown"""
    
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        title = kwargs.get('title', 'DataCraft Export')
        max_rows = kwargs.get('max_rows', 100)
        
        # Limita número de linhas
        df_to_export = df if max_rows is None else df.head(max_rows)
        
        # Gera Markdown
        md_content = f"# {title}\n\n"
        md_content += f"**Total de registros:** {len(df)}  \n"
        md_content += f"**Total de colunas:** {len(df.columns)}  \n\n"
        
        # Tabela
        md_content += "## Dados\n\n"
        md_content += df_to_export.to_markdown(index=False)
        
        md_content += f"\n\n---\n*Gerado por DataCraft em {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}*"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(md_content)
        
        return output_path
    
    def get_supported_formats(self) -> list:
        return ['markdown', 'md']