from io import StringIO
import pandas as pd
from pathlib import Path
import PySimpleGUI as sg
import subprocess
import sys

_encoding = 'CP1250'

def read_postbank_tx_csv(filename):
    csv_text = ''
    with open(filename, 'r', encoding=_encoding) as f:
        line = f.readline()
        while line:
            if line.strip().startswith('gebuchte Ums'): # gebuchte Umsätze
                csv_text = f.read()
                break
            line = f.readline()
    if len(csv_text) == 0: # Must have been a pure CSV instead
        with open(filename, 'r', encoding=_encoding) as f:
            csv_text = f.read()
    df = pd.read_csv(StringIO(csv_text), sep=';', decimal=',', thousands='.')
    last_column_has_value = (~df[df.columns[-1]].isna()).any() # Any value is not NaN
    if not last_column_has_value:
        df.drop(df.columns[-1], axis=1, inplace=True)
    return df

def join_unique_and_sort_dfs(dfs):
    df = pd.concat(dfs)
    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.sort_values(by=['Buchungsdatum', 'Wertstellung'], inplace=True)
    return df

def format_date_in_place_DE(df, column_names):
    for column_name in column_names:
        df[column_name] = pd.to_datetime(df[column_name].astype(str), format='%d%m%Y')
        df[column_name] = df[column_name].dt.strftime('%d.%m.%Y')

def main():
    file_types = (("(Postbank) CSV", "*.csv"),)
    layout = [
        [sg.Text('Datei A'), sg.Input(), sg.FileBrowse(file_types=file_types)],
        [sg.Text('Datei B'), sg.Input(), sg.FileBrowse(file_types=file_types)],
        [sg.Input(), sg.FileSaveAs(file_types=file_types)],
        [sg.OK(), sg.Cancel()]
    ]

    window = sg.Window('Postbank CSV Merger (keine Gewähr)', layout)
    event, values = window.read()
    window.close()

    if event == sg.WIN_CLOSED or event == 'Cancel':
        sys.exit()

    file_a = values[0].strip()
    file_b = values[1].strip()
    file_out = values[2].strip()

    if not file_a or not file_b:
        sg.Popup('Bitte zwei Dateien auswählen')
        sys.exit()

    if not file_out:
        sg.Popup('Bitte eine Ausgabedatei auswählen')
        sys.exit()

    try:
        df_a = read_postbank_tx_csv(file_a)
    except:
        sg.Popup('Fehler beim Lesen der Datei A')
        sys.exit()

    try:
        df_b = read_postbank_tx_csv(file_b)
    except:
        sg.Popup('Fehler beim Lesen der Datei B')
        sys.exit()

    try:
        df_out = join_unique_and_sort_dfs([df_a, df_b])
    except:
        sg.Popup('Fehler beim Zusammenführen der Dateien')
        sys.exit()

    try:
        format_date_in_place_DE(df_out, ['Buchungsdatum', 'Wertstellung'])
        df_out.to_csv(file_out, sep=';', decimal=',', index=False, encoding=_encoding)
    except:
        sg.Popup('Fehler beim Speichern der Ausgabe')
        sys.exit()

    layout = [
        [sg.Text(f'Datei wurde gespeicher:\n{file_out}\n\n')],
        [sg.Text('Ordner öffnen?')],
        [sg.OK(), sg.Cancel()]
    ]
    window = sg.Window('Postbank CSV Merger (keine Gewähr)', layout)
    event, _ = window.read()
    window.close()

    if event == 'OK':
        subprocess.run(['explorer', Path(file_out).parent])

main()
