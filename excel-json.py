"""
Excel to CSV/JSON Converter
Converte arquivos .xlsx para CSV (com limite de linhas) e JSON
"""

import os
import math
import glob
import pandas as pd
from openpyxl import load_workbook

def setup_directories():
    """Cria as pastas necessárias"""
    os.makedirs("files", exist_ok=True)
    os.makedirs("output/csv", exist_ok=True)
    os.makedirs("output/json", exist_ok=True)

def load_excel_files(folder="files"):
    """Carrega todos os arquivos .xlsx da pasta"""
    paths = glob.glob(f"{folder}/*.xlsx")
    dataframes = {}
    
    for path in paths:
        try:
            df = pd.read_excel(path)
            nome = os.path.splitext(os.path.basename(path))[0]
            dataframes[nome] = df
            print(f"✓ {nome} carregado: {path}")
        except Exception as e:
            print(f"✗ Erro ao ler {path}: {e}")
    
    return dataframes

def convert_to_csv(dataframes, limite_de_linhas=20000):
    """Converte para CSV com limite de linhas"""
    for nome, df in dataframes.items():
        total_de_linhas = len(df)
        total_de_arquivos = math.ceil(total_de_linhas / limite_de_linhas)
        
        for i in range(total_de_arquivos):
            inicio = i * limite_de_linhas
            fim = inicio + limite_de_linhas
            selecao = df.iloc[inicio:fim]
            
            selecao.to_csv(f"output/csv/{nome}(pt{i+1}).csv", index=False)
            print(f"✓ {nome}(pt{i+1}).csv gerado: linhas {inicio}-{min(fim, total_de_linhas)}")

def convert_to_json(dataframes):
    """Converte para JSON"""
    for nome, df in dataframes.items():
        df.to_json(f"output/json/{nome}.json", orient="records", indent=2)
        print(f"✓ {nome}.json gerado com sucesso!")

def main():
    print("=== Excel to CSV/JSON Converter ===\n")
    
    setup_directories()
    print("\n1. Carregando arquivos Excel...")
    dataframes = load_excel_files()
    
    if not dataframes:
        print("\n⚠️  Nenhum arquivo .xlsx encontrado na pasta 'files/'")
        return
    
    print(f"\n2. Convertendo para CSV (limite: 20.000 linhas)...")
    convert_to_csv(dataframes)
    
    print(f"\n3. Convertendo para JSON...")
    convert_to_json(dataframes)
    
    print("\n✅ Conversão concluída!")

if __name__ == "__main__":
    main()