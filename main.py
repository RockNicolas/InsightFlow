import pandas as pd
import os

# 1. Configurações Iniciais
arquivo_excel = "Controle Km e Horímetro_março26.xlsx"
aba_desejada = "01 A 08.03"  # Nome exato da aba dentro do Excel

def gerar_relatorio_semanal(caminho, aba):
    if not os.path.exists(caminho):
        print(f"Erro: O arquivo '{caminho}' não foi encontrado na pasta.")
        return

    try:
        # Lendo o Excel 
        # header=4 diz ao pandas que o cabeçalho real (MÁQUINA, OPERADOR...) está na linha 5
        df = pd.read_excel(caminho, sheet_name=aba, header=4)

        # Limpeza: remove linhas onde a coluna 'MÁQUINA' está vazia
        df = df.dropna(subset=[df.columns[1]])

        print(f"\n{'='*60}")
        print(f"RELATÓRIO SEMANAL: {aba}")
        print(f"{'='*60}")
        print(f"{'MÁQUINA':<15} | {'OPERADOR':<20} | {'TOTAL HORAS':<10}")
        print("-" * 60)

        total_da_semana = 0

        for index, row in df.iterrows():
            maquina = str(row.iloc[1])    # Coluna B (MÁQUINA)
            operador = str(row.iloc[4])   # Coluna E (OPERADOR)
            # A última coluna é o 'TOTAL HORAS TRABALHADAS'
            total_horas = row.iloc[-1] 
            
            try:
                # Trata o valor para garantir que seja um número somável
                valor = float(total_horas) if pd.notnull(total_horas) else 0.0
            except:
                valor = 0.0

            if valor != 0:
                print(f"{maquina:<15} | {operador:<20} | {valor:<10.2f}")
                total_da_semana += valor

        print("-" * 60)
        print(f"{'TOTAL GERAL DA SEMANA:':<38} {total_da_semana:.2f} Horas")
        print(f"{'='*60}\n")

    except Exception as e:
        print(f"Erro ao ler a aba '{aba}': {e}")
        print("Verifique se o nome da aba está correto no Excel.")

# Executar
gerar_relatorio_semanal(arquivo_excel, aba_desejada)