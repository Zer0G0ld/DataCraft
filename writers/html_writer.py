"""
Escritor para arquivos HTML
"""

import pandas as pd
from .base_writer import BaseWriter

class HTMLWriter(BaseWriter):
    """Escritor para formato HTML"""
    
    def write(self, df: pd.DataFrame, output_path: str, **kwargs) -> str:
        title = kwargs.get('title', 'DataCraft Export')
        max_rows = kwargs.get('max_rows', None)
        
        # Limita número de linhas se necessário
        df_to_export = df if max_rows is None else df.head(max_rows)
        
        # Gera HTML
        html_content = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>{title}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            max-width: 100%;
            overflow-x: auto;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th {{
            background-color: #2c3e50;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
        }}
        td {{
            padding: 8px 12px;
            border-bottom: 1px solid #ddd;
        }}
        tr:hover {{
            background-color: #f5f5f5;
        }}
        .info {{
            margin-bottom: 20px;
            padding: 10px;
            background-color: #e8f4f8;
            border-radius: 4px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="info">
            <strong>📊 Total de registros:</strong> {len(df)} | 
            <strong>📋 Total de colunas:</strong> {len(df.columns)}
        </div>
        {df_to_export.to_html(index=False, classes='data-table')}
    </div>
</body>
</html>
"""
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        return output_path
    
    def get_supported_formats(self) -> list:
        return ['html', 'htm']