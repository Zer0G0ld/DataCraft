#!/usr/bin/env python3
"""
DataCraft - Conversor Universal de Dados
Transforme qualquer formato em qualquer formato!
"""

import sys
import os

# Adiciona o diretório atual ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import main

if __name__ == "__main__":
    main()