import pandas as pd
import os
import unicodedata


def _find_sheet_name(file_path, requested_sheet):
    requested = str(requested_sheet or "").strip().strip("'").strip()
    if not requested:
        return requested_sheet

    try:
        xl = pd.ExcelFile(file_path)
        for name in xl.sheet_names:
            candidate = str(name or "").strip()
            if candidate.strip("'").strip() == requested:
                return name
    except Exception:
        pass

    return requested_sheet


def _normalize_text(value):
    """Normaliza texto removendo acentos e padronizando para comparacoes."""
    text = str(value or "").strip().upper()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _detect_header_row(file_path, sheet):
    """
    Detecta a linha de cabecalho da aba.
    Algumas semanas trazem o cabecalho uma linha acima/abaixo.
    """
    preview = pd.read_excel(file_path, sheet_name=sheet, header=None, nrows=15)

    for idx in range(len(preview.index)):
        row_values = [_normalize_text(v) for v in preview.iloc[idx].tolist()]
        has_machine = any("MAQUINA" in v for v in row_values)
        has_operator = any("OPERADOR" in v for v in row_values)
        if has_machine and has_operator:
            return idx

    # fallback historico do projeto
    return 4

def _to_float_or_none(value):
    """Converte valor para float quando possível."""
    try:
        if pd.isnull(value):
            return None
        return float(value)
    except (TypeError, ValueError):
        return None

def get_weekly_data(file_path, sheet):
    """Extrai dados de todas as frotas e concatena Máquina + Placa."""
    sheet = _find_sheet_name(file_path, str(sheet or "").strip())
    try:
        header_row = _detect_header_row(file_path, sheet)
        df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)

        df.columns = [str(col).strip().replace('\n', ' ') for col in df.columns]
        normalized_cols = [_normalize_text(col) for col in df.columns]

        processed_data = []
    
        hour_cols = [i for i, col in enumerate(df.columns) 
                    if "HORAS TRABALHADAS" in normalized_cols[i] and "TOTAL" not in normalized_cols[i]]
        final_cols = [i for i, col in enumerate(df.columns)
                    if "FINAL" in normalized_cols[i] and "HOR" in normalized_cols[i]]

        for _, row in df.iterrows():
            machine_name = row.iloc[1]  
            plate = row.iloc[2]         
            operator = row.iloc[4]      
            
            if pd.notnull(machine_name) and str(machine_name).strip() != "" and str(machine_name).upper() != "MÁQUINA":
                
                full_machine_name = f"{str(machine_name).strip()} {str(plate).strip() if pd.notnull(plate) else ''}".strip()
                
                weekly_sum = 0
                has_data = False
                final_meter = None
                
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

                for idx in final_cols:
                    if idx < len(row):
                        parsed = _to_float_or_none(row.iloc[idx])
                        if parsed is not None and parsed >= 0:
                            final_meter = parsed
                
                processed_data.append({
                    'machine': full_machine_name,
                    'operator': str(operator).strip() if pd.notnull(operator) else "N/A",
                    'hours': weekly_sum,
                    'final_meter': final_meter
                })
        
        return processed_data

    except Exception as e:
        print(f"Erro no Processador: {e}")
        return []