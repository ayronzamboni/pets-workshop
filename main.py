import os
import sys
import io
import pandas as pd
import PySimpleGUI as sg
from pathlib import Path
from datetime import datetime
import traceback

LOGFILE = "app_edit_planilha.log"
BACKUP_DIR = "backups"
MAX_TABLE_ROWS = 1000

def log(msg):
    with open(LOGFILE, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().isoformat()} - {msg}\n")

def create_backup(file_path):
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        base = Path(file_path)
        if base.exists():
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = Path(BACKUP_DIR) / f"{base.stem}_backup_{stamp}{base.suffix}"
            with open(base, "rb") as src, open(dest, "wb") as dst:
                dst.write(src.read())
            log(f"Backup criado: {dest}")
    except Exception as e:
        log(f"Erro criando backup: {e}")

def read_file(path):
    path = Path(path)
    if path.suffix.lower() in [".xls", ".xlsx"]:
        try:
            xls = pd.read_excel(path, sheet_name=None, engine="openpyxl")
            for k, v in xls.items():
                if not isinstance(v, pd.DataFrame):
                    xls[k] = pd.DataFrame(v)
            return xls
        except Exception as e:
            log(f"Erro lendo xlsx: {e}")
            raise
    elif path.suffix.lower() in [".csv", ".txt"]:
        df = pd.read_csv(path, encoding="utf-8", sep=None, engine="python")
        return {"Sheet1": df}
    else:
        raise ValueError("Formato não suportado. Use .xlsx ou .csv")

def write_file(path, sheets_dict):
    path = Path(path)
    if path.suffix.lower() in [".xls", ".xlsx"]:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            for name, df in sheets_dict.items():
                df.to_excel(writer, sheet_name=str(name)[:31], index=False)
    elif path.suffix.lower() in [".csv", ".txt"]:
        list(sheets_dict.values())[0].to_csv(path, index=False, encoding="utf-8")
    else:
        raise ValueError("Formato não suportado para escrita.")

def df_to_table_values(df, limit=MAX_TABLE_ROWS):
    df_show = df.copy()
    df_show = df_show.fillna("").astype(str)
    if len(df_show) > limit:
        return df_show.head(limit).values.tolist(), True
    return df_show.values.tolist(), False

def edit_row_popup(column_names, row_values, title="Editar linha"):
    layout = []
    for i, col in enumerate(column_names):
        layout.append([sg.Text(str(col), size=(20,1)), sg.Input(default_text=str(row_values[i]), key=f"-IN-{i}-", expand_x=True)])
    layout.append([sg.Button("Salvar"), sg.Button("Cancelar")])
    win = sg.Window(title, layout, modal=True, resizable=True)
    event, vals = win.read()
    result = None
    if event == "Salvar":
        result = [vals[f"-IN-{i}-"] for i in range(len(column_names))]
    win.close()
    return result

def add_row_popup(column_names, title="Adicionar linha"):
    layout = []
    for i, col in enumerate(column_names):
        layout.append([sg.Text(str(col), size=(20,1)), sg.Input(key=f"-IN-{i}-", expand_x=True)])
    layout.append([sg.Button("Adicionar"), sg.Button("Cancelar")])
    win = sg.Window(title, layout, modal=True, resizable=True)
    event, vals = win.read()
    result = None
    if event == "Adicionar":
        result = [vals[f"-IN-{i}-"] for i in range(len(column_names))]
    win.close()
    return result

def main():
    sg.theme("LightBlue")

    file_path = ""
    sheets = {}
    current_sheet = None

    layout_top = [
        [sg.Button("Abrir arquivo", key="-OPEN-"), sg.Text("", key="-FILE-")],
        [sg.Text("Aba:"), sg.Combo([], key="-SHEET-", size=(30,1), enable_events=True),
         sg.Button("Atualizar tabela"), sg.Text("Buscar:", pad=((10,0),(0,0))), sg.Input(key="-SEARCH-"), sg.Button("Filtrar")]
    ]

    table = sg.Table(values=[[]], headings=[], key="-TABLE-", enable_events=True,
                     display_row_numbers=True, auto_size_columns=False, col_widths=[20]*8,
                     justification="left", num_rows=20, select_mode=sg.TABLE_SELECT_MODE_BROWSE)

    layout_bottom = [
        [sg.Button("Editar linha"), sg.Button("Adicionar linha"), sg.Button("Excluir seleção"), sg.Button("Exportar CSV"),
         sg.Button("Salvar"), sg.Button("Salvar como..."), sg.Text("", key="-STATUS-", size=(40,1))]
    ]

    layout = [
        [sg.Column(layout_top, expand_x=True)],
        [table],
        [sg.Column(layout_bottom)]
    ]

    window = sg.Window("Editor de Planilha Simples", layout, resizable=True, finalize=True)
    df_filtered = None

    while True:
        event, values = window.read(timeout=100)
        if event == sg.WIN_CLOSED:
            break

        if event == "-OPEN-":
            path = sg.popup_get_file("Escolha um arquivo .xlsx ou .csv", file_types=(("Excel Files","*.xlsx;*.xls"),("CSV Files","*.csv;*.txt")), no_window=True)
            if path:
                try:
                    sheets = read_file(path)
                    file_path = path
                    sheet_names = list(sheets.keys())
                    window["-SHEET-"].update(values=sheet_names, value=sheet_names[0] if sheet_names else "")
                    current_sheet = sheet_names[0] if sheet_names else None
                    window["-FILE-"].update(Path(path).name)
                    df_filtered = None
                    window["-STATUS-"].update(f"Arquivo carregado: {Path(path).name}")
                    log(f"Arquivo aberto: {path}")
                    if current_sheet:
                        df = sheets[current_sheet]
                        vals, _ = df_to_table_values(df)
                        window["-TABLE-"].update(values=vals, headings=list(df.columns))
                except Exception as e:
                    sg.popup_error("Erro ao abrir arquivo:", str(e))
                    log(f"Erro open file: {traceback.format_exc()}")

        if event == "-SHEET-":
            current_sheet = values["-SHEET-"]
            df_filtered = None
            if current_sheet:
                df = sheets[current_sheet]
                vals, _ = df_to_table_values(df)
                window["-TABLE-"].update(values=vals, headings=list(df.columns))
                window["-STATUS-"].update(f"Aba: {current_sheet}")

        if event == "Filtrar":
            q = values["-SEARCH-"].strip()
            if current_sheet and q != "":
                df = sheets[current_sheet]
                mask = df.astype(str).apply(lambda row: row.str.contains(q, case=False, na=False)).any(axis=1)
                df_filtered = df[mask]
                vals, _ = df_to_table_values(df_filtered)
                window["-TABLE-"].update(values=vals, headings=list(df_filtered.columns))
                window["-STATUS-"].update(f"Filtro '{q}' aplicado: {len(df_filtered)} resultados")
            else:
                df_filtered = None
                if current_sheet:
                    df = sheets[current_sheet]
                    vals, _ = df_to_table_values(df)
                    window["-TABLE-"].update(values=vals, headings=list(df.columns))
                    window["-STATUS-"].update("Filtro limpo")

        if event == "Atualizar tabela":
            if current_sheet:
                df = sheets[current_sheet]
                vals, _ = df_to_table_values(df)
                window["-TABLE-"].update(values=vals, headings=list(df.columns))

        if event == "Editar linha":
            sel = values["-TABLE-"]
            if not sel:
                sg.popup("Selecione uma linha na tabela para editar.")
                continue
            row_index = sel[0]
            df = df_filtered if df_filtered is not None else sheets[current_sheet]
            real_index = df.index[row_index]
            row = df.loc[real_index].tolist()
            col_names = list(df.columns)
            new_row = edit_row_popup(col_names, row)
            if new_row is not None:
                try:
                    for i, col in enumerate(col_names):
                        sheets[current_sheet].at[real_index, col] = try_parse_value(new_row[i], sheets[current_sheet].dtypes.get(col))
                    vals, _ = df_to_table_values(sheets[current_sheet])
                    window["-TABLE-"].update(values=vals, headings=list(sheets[current_sheet].columns))
                    window["-STATUS-"].update("Linha editada")
                except Exception as e:
                    sg.popup_error("Erro ao salvar edição:", e)

        if event == "Adicionar linha":
            if not current_sheet:
                sg.popup("Abra um arquivo e selecione uma aba primeiro.")
                continue
            df = sheets[current_sheet]
            col_names = list(df.columns)
            new_row = add_row_popup(col_names)
            if new_row is not None:
                try:
                    new_series = {}
                    for i, col in enumerate(col_names):
                        new_series[col] = try_parse_value(new_row[i], df.dtypes.get(col))
                    sheets[current_sheet] = df._append(new_series, ignore_index=True)
                    vals, _ = df_to_table_values(sheets[current_sheet])
                    window["-TABLE-"].update(values=vals, headings=list(sheets[current_sheet].columns))
                    window["-STATUS-"].update("Linha adicionada")
                except Exception as e:
                    sg.popup_error("Erro ao adicionar linha:", e)

        if event == "Excluir seleção":
            sel = values["-TABLE-"]
            if not sel:
                sg.popup("Selecione uma linha para excluir.")
                continue
            confirm = sg.popup_yes_no("Confirma excluir a linha selecionada?")
            if confirm == "Yes":
                row_index = sel[0]
                df = df_filtered if df_filtered is not None else sheets[current_sheet]
                real_index = df.index[row_index]
                try:
                    sheets[current_sheet] = sheets[current_sheet].drop(real_index).reset_index(drop=True)
                    df_filtered = None
                    vals, _ = df_to_table_values(sheets[current_sheet])
                    window["-TABLE-"].update(values=vals, headings=list(sheets[current_sheet].columns))
                    window["-STATUS-"].update("Linha excluída")
                except Exception as e:
                    sg.popup_error("Erro ao excluir:", e)

        if event == "Exportar CSV":
            if not current_sheet:
                sg.popup("Abra um arquivo e selecione uma aba primeiro.")
                continue
            out = sg.popup_get_file("Salvar CSV como...", save_as=True, file_types=(("CSV Files", "*.csv"),), default_extension=".csv", no_window=True)
            if out:
                try:
                    sheets[current_sheet].to_csv(out, index=False, encoding="utf-8")
                    sg.popup("Exportado com sucesso.")
                except Exception as e:
                    sg.popup_error("Erro ao exportar CSV:", e)

        if event == "Salvar":
            if not file_path:
                sg.popup("Abra um arquivo primeiro.")
                continue
            try:
                create_backup(file_path)
                write_file(file_path, sheets)
                sg.popup("Salvo com sucesso.")
                window["-STATUS-"].update(f"Salvo: {Path(file_path).name}")
            except Exception as e:
                sg.popup_error("Erro ao salvar:", e)

        if event == "Salvar como...":
            if not sheets:
                sg.popup("Abra um arquivo primeiro.")
                continue
            out = sg.popup_get_file("Salvar como...", save_as=True, file_types=(("Excel Files", "*.xlsx"),("CSV Files","*.csv")), default_extension=".xlsx", no_window=True)
            if out:
                try:
                    write_file(out, sheets)
                    sg.popup("Salvo com sucesso.")
                except Exception as e:
                    sg.popup_error("Erro ao salvar como:", e)

    window.close()

def try_parse_value(val, dtype):
    if val is None:
        return val
    s = str(val).strip()
    if s == "":
        return ""
    try:
        if dtype is not None:
            if "int" in str(dtype):
                return int(float(s.replace(",",".").replace("R$","").replace(".","").replace(" ","")))
            if "float" in str(dtype):
                return float(s.replace("R$","").replace(".","").replace(",",".").strip())
    except:
        pass
    try:
        s2 = s.replace("R$","").replace(" ","")
        if "," in s2 and "." in s2:
            return float(s2.replace(".","").replace(",","."))
        if "," in s2 and s2.count(",") == 1 and all(c.isdigit() or c in ",-" for c in s2):
            return float(s2.replace(",","."))
        if all(c.isdigit() or c in ".-" for c in s2):
            return float(s2)
    except:
        pass
    return val

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        sg.popup_error("Erro inesperado:", e)
        log(f"Crash: {traceback.format_exc()}")
