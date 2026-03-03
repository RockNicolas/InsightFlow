import pandas as pd
import os

def get_weekly_data(file_path, sheet):
    """Extrai dados de todas as frotas e concatena Máquina + Placa."""
    try:
    
        df = pd.read_excel(file_path, sheet_name=sheet, header=4)

        df.columns = [str(col).strip().replace('\n', ' ') for col in df.columns]

        processed_data = []
    
        hour_cols = [i for i, col in enumerate(df.columns) 
                    if "HORAS TRABALHADAS" in col.upper() and "TOTAL" not in col.upper()]

        for _, row in df.iterrows():
            machine_name = row.iloc[1]  
            plate = row.iloc[2]         
            operator = row.iloc[4]      
            
            if pd.notnull(machine_name) and str(machine_name).strip() != "" and str(machine_name).upper() != "MÁQUINA":
                
                full_machine_name = f"{str(machine_name).strip()} {str(plate).strip() if pd.notnull(plate) else ''}".strip()
                
                weekly_sum = 0
                has_data = False
                
                for idx in hour_cols:
                    if idx < len(row):
                        val = row.iloc[idx]
                        try:
                            f_val = float(val)
                            
                            if 0 <= f_val < 1000:
                                weekly_sum += f_val
                                has_data = True
                        except:
                            continue
                
                processed_data.append({
                    'machine': full_machine_name,
                    'operator': str(operator).strip() if pd.notnull(operator) else "N/A",
                    'hours': weekly_sum
                })
        
        return processed_data

    except Exception as e:
        print(f"Erro no Processador: {e}")
        return []