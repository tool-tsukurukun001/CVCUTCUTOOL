"""CVCUTCUTOOL – simple GUI wrapper

Requirements:
    pip install PySimpleGUI openpyxl pydub faster-whisper fuzzywuzzy python-Levenshtein

Run:
    python cvcutcutoool_gui.py
"""
import PySimpleGUI as sg
from pathlib import Path
from utils import load_script_list
from recognizer import transcribe_full
from matcher import match_segments
from naming import generate_filename
from pydub import AudioSegment
import threading


def choose_excel():
    path = sg.popup_get_file("CVリスト(Excel) を選択", file_types=(("Excel","*.xlsx"),))
    return Path(path) if path else None

def choose_folder(title):
    path = sg.popup_get_folder(title)
    return Path(path) if path else None

def worker(values, window):
    cv_path: Path = values['CV_PATH']
    sheet     = values['SHEET'] or 'Sheet1'
    no_cell   = values['NO_CELL'] or 'A'
    text_cell = values['TEXT_CELL'] or 'B'
    wav_dir: Path = values['WAV_DIR']
    out_prefix  = values['PREFIX'] or 'CV_'
    out_root = Path.cwd() / 'output'
    out_root.mkdir(exist_ok=True)

    # load script
    scripts = load_script_list(cv_path, sheet_name=sheet)

    # table clear
    window['TABLE'].update(values=[])

    for wav in wav_dir.glob('*.wav'):
        wav_sub = out_root / wav.stem
        wav_sub.mkdir(exist_ok=True)

        segments = transcribe_full(str(wav))
        matches  = match_segments(segments, scripts, threshold=80)
        audio    = AudioSegment.from_file(wav)

        table_rows = []
        for idx, m in enumerate(matches, start=1):
            start_ms = int(m['start']*1000)
            end_ms   = int(m['end']*1000)
            clip = audio[start_ms:end_ms]
            fname = f"{out_prefix}{int(m['script_number']):03d}_{idx:02d}.wav"
            clip.export(wav_sub / fname, format='wav')
            table_rows.append([m['script_number'], m['text'], fname])
        window['TABLE'].update(values=table_rows)

    sg.popup_ok('完了しました')


def main():
    sg.theme('LightBlue')

    left_layout = [
        [sg.Text('CVリスト'), sg.Input(key='CV_PATH', size=(40,1), readonly=True), sg.FileBrowse('選択', file_types=(('Excel','*.xlsx'),))],
        [sg.Text('シート'), sg.Input(key='SHEET', size=(10,1))],
        [sg.Text('Noセル'), sg.Input(key='NO_CELL', size=(10,1))],
        [sg.Text('台詞セル'), sg.Input(key='TEXT_CELL', size=(10,1))],
        [sg.Text('wavフォルダ'), sg.Input(key='WAV_DIR', size=(40,1), readonly=True), sg.FolderBrowse('選択')],
        [sg.Text('出力ファイル名 プレフィックス'), sg.Input(key='PREFIX', size=(25,1))],
        [sg.Button('WAV出力', key='RUN', size=(20,2), button_color=('white','blue'))]
    ]

    table_head = ['No','台詞','出力wav名']
    right_layout = [[sg.Table(values=[], headings=table_head, key='TABLE', auto_size_columns=True, col_widths=[6,40,25], num_rows=25, justification='left')]]

    layout = [[sg.Column(left_layout), sg.VSeparator(), sg.Column(right_layout)]]

    window = sg.Window('CVCUTCUTOOOON Ver1.0.0', layout, finalize=True, resizable=True)

    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, 'Exit'):
            break
        if event == 'RUN':
            if not values['CV_PATH'] or not values['WAV_DIR']:
                sg.popup_error('CVリストとwavフォルダを選択してください')
                continue
            threading.Thread(target=worker, args=(values, window), daemon=True).start()

    window.close()

if __name__ == '__main__':
    main()
