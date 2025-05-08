import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from utils import load_script_list
from recognizer import transcribe_full
from matcher import match_segments
from pydub import AudioSegment
from pydub.playback import play
from fuzzywuzzy import fuzz
import threading
from splitter import split_audio
import os
import csv
import json
import tempfile
import uuid
import simpleaudio as sa
import wave
import time
import sounddevice as sd
import numpy as np


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CVCUTCUTOOL')
        self.geometry('1050x650')

        # 再生用の変数
        self.current_audio = None
        self.is_playing = False
        self.play_thread = None

        # ---------- 左ペイン ----------
        frm = tk.Frame(self)
        frm.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        # CVリスト (Excel)
        self.cv_path = tk.StringVar()
        tk.Label(frm, text='CVリスト').grid(row=0, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.cv_path, width=42).grid(row=0, column=1)
        tk.Button(frm, text='参照', command=self.browse_cv).grid(row=0, column=2)

        # シート名
        self.sheet = tk.StringVar(value='Sheet1')
        tk.Label(frm, text='シート').grid(row=1, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.sheet, width=12).grid(row=1, column=1, sticky='w')

        # No列 / 台詞列
        self.no_col = tk.StringVar(value='A')
        self.text_col = tk.StringVar(value='B')
        tk.Label(frm, text='No列').grid(row=2, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.no_col, width=5).grid(row=2, column=1, sticky='w')
        tk.Label(frm, text='台詞列').grid(row=3, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.text_col, width=5).grid(row=3, column=1, sticky='w')

        # WAVフォルダ
        self.wav_dir = tk.StringVar()
        tk.Label(frm, text='wavフォルダ').grid(row=4, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.wav_dir, width=42).grid(row=4, column=1)
        tk.Button(frm, text='参照', command=self.browse_wav).grid(row=4, column=2)

        # 開始行
        self.start_row = tk.StringVar(value='2')
        tk.Label(frm, text='開始行').grid(row=5, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.start_row, width=8).grid(row=5, column=1, sticky='w')

        # 出力プレフィックス
        self.prefix = tk.StringVar(value='CVCUTCUTOOL_')
        tk.Label(frm, text='出力ファイル名プレフィックス').grid(row=6, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.prefix, width=25).grid(row=6, column=1, sticky='w')

        # 出力先フォルダ
        self.out_dir = tk.StringVar()
        tk.Label(frm, text='出力先フォルダ').grid(row=7, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.out_dir, width=42).grid(row=7, column=1)
        tk.Button(frm, text='参照', command=self.browse_out).grid(row=7, column=2)

        # スクリプト照合しきい値
        self.threshold = tk.IntVar(value=80)
        tk.Scale(
            frm, from_=0, to=100, orient='horizontal', variable=self.threshold,
            label='しきい値 (ゆるい 0〜100 きびしい)', length=250
        ).grid(row=8, column=0, columnspan=3, pady=5)

        # 音量しきい値 (dBFS)
        self.volume_thresh = tk.IntVar(value=-40)
        tk.Scale(
            frm, from_=-60, to=0, orient='horizontal', variable=self.volume_thresh,
            label='音量しきい値 (dBFS)', length=250
        ).grid(row=9, column=0, columnspan=3, pady=5)

        # WAV出力ボタン
        tk.Button(
            frm, text='WAV 出力', bg='royalblue', fg='white', width=22,
            command=self.run
        ).grid(row=10, column=0, columnspan=3, pady=15)

        # --- ここから下にコントロールボタンを縦に配置 ---
        self.play_button = tk.Button(frm, text='再生', command=self.play_selected, width=22)
        self.play_button.grid(row=11, column=0, columnspan=3, pady=2)
        self.continuous_play_button = tk.Button(frm, text='連続再生', command=self.play_continuous, width=22)
        self.continuous_play_button.grid(row=12, column=0, columnspan=3, pady=2)
        self.stop_button = tk.Button(frm, text='停止', command=self.stop_playback, width=22)
        self.stop_button.grid(row=13, column=0, columnspan=3, pady=2)
        self.csv_button = tk.Button(frm, text='CSV出力', command=self.export_csv, width=22)
        self.csv_button.grid(row=14, column=0, columnspan=3, pady=10)

        # ---------- 右ペイン：結果テーブル ----------
        cols = ('No', '台詞', '出力wav名')
        self.table = ttk.Treeview(self, columns=cols, show='headings', height=28)
        for c in cols:
            self.table.heading(c, text=c)
            self.table.column(c, width=260 if c == '台詞' else 140, anchor='w')
        self.table.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # テーブルの選択イベントをバインド
        self.table.bind('<<TreeviewSelect>>', self.on_select)
        self.table.bind('<Double-1>', self.on_double_click)
        
        # 編集用の変数
        self.edit_item = None
        self.edit_column = None
        self.edit_entry = None
        
        # メニューバーの作成
        self.menu_bar = tk.Menu(self)
        self.config(menu=self.menu_bar)
        
        # ファイルメニュー
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label='ファイル', menu=file_menu)
        file_menu.add_command(label='プロジェクトを保存', command=self.save_project)
        file_menu.add_command(label='プロジェクトを開く', command=self.load_project)
        file_menu.add_separator()
        file_menu.add_command(label='終了', command=self.quit)

    # ダイアログ: ファイル/フォルダ選択
    def browse_cv(self):
        p = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')])
        if p:
            self.cv_path.set(p)

    def browse_wav(self):
        p = filedialog.askdirectory()
        if p:
            self.wav_dir.set(p)

    def browse_out(self):
        p = filedialog.askdirectory()
        if p:
            self.out_dir.set(p)

    # 実行トリガー
    def run(self):
        if not self.cv_path.get() or not self.wav_dir.get():
            messagebox.showerror('エラー', 'CVリストとwavフォルダを指定してください')
            return
        # 進行中インジケーター
        self.config(cursor='watch')
        if not hasattr(self, 'progress_label'):
            self.progress_label = tk.Label(self, text='処理中...', fg='red', font=('Arial', 16))
            self.progress_label.pack(side=tk.TOP, pady=5)
        else:
            self.progress_label.config(text='処理中...')
        self.update()
        threading.Thread(target=self.process, daemon=True).start()

    # mis 用の最良スクリプト推定
    def best_script(self, text, scripts):
        return max(scripts, key=lambda s: fuzz.partial_ratio(text, s['text']))['no']

    # メイン処理
    def process(self):
        self.table.delete(*self.table.get_children())

        # スクリプト読み込み
        scripts = load_script_list(
            Path(self.cv_path.get()),
        sheet_name=self.sheet.get(),
        start_row=int(self.start_row.get() or 2),
        no_col=self.no_col.get(),        # ← 追加
        text_col=self.text_col.get()     # ← 追加
    )

        wav_dir = Path(self.wav_dir.get())
        out_root = Path(self.out_dir.get()) if self.out_dir.get() else Path.cwd() / 'output'
        out_root.mkdir(exist_ok=True)

        for wav in wav_dir.glob('*.wav'):
            wav_sub = out_root / wav.stem
            wav_sub.mkdir(exist_ok=True)

            # 無音区間で分割（パラメータ調整）
            chunk_segments = split_audio(
                str(wav),
                min_silence_len=700,   # 無音とみなす最小長さ（ms）
                silence_thresh=-40,    # 無音判定の音量（dBFS）
                keep_silence=400      # 前後に保持する無音長（ms）
            )
            segments = []
            for idx, chunk in enumerate(chunk_segments, start=1):
                # チャンクを一時ファイルとして保存
                tmp_chunk_path = wav_sub / f"_tmp_chunk_{idx:03d}.wav"
                chunk['audio'].export(tmp_chunk_path, format='wav')
                # Whisperで認識
                chunk_result = transcribe_full(str(tmp_chunk_path))
                # 認識結果を格納（タイムスタンプはチャンクの相対値なので、絶対値に変換）
                for seg in chunk_result:
                    seg_abs = {
                        'start': chunk['start'] + seg['start'],
                        'end': chunk['start'] + seg['end'],
                        'text': seg['text']
                    }
                    segments.append(seg_abs)
                tmp_chunk_path.unlink()  # 一時ファイル削除

            # デバッグ: 認識結果を全て出力
            print(f"[DEBUG] segments for {wav.name}")
            for seg in segments:
                print(f"  {seg['start']:.2f}-{seg['end']:.2f}: {seg['text']}")

            # 2) 音量フィルタ
            vol_th = self.volume_thresh.get()
            filtered = []
            audio = AudioSegment.from_file(wav)
            for seg in segments:
                clip = audio[int(seg['start']*1000): int(seg['end']*1000)]
                if clip.dBFS >= vol_th:
                    filtered.append(seg)
            segments = filtered

            # 3) スクリプト照合
            matches = match_segments(segments, scripts, threshold=self.threshold.get())
            matched_keys = {(m['start'], m['end']) for m in matches}

            count_dict = {}
            mis_count_dict = {}
            # 正常マッチ・mis両方対応
            for m in matches:
                if m.get('is_mis') and m['script_number'] == '000':
                    # misファイル（日本語・英語以外）は000_00_misXX.wav
                    fname = f"{self.prefix.get()}000_00_mis{m['mis_index']:02d}.wav"
                    text_disp = m['text']
                    no_disp = '000(mis)'
                else:
                    num = int(m['script_number'])
                    if m.get('is_mis'):
                        mis_count_dict[num] = mis_count_dict.get(num, 0) + 1
                        ver = count_dict.get(num, 0) + 1
                        mis_ver = mis_count_dict[num]
                        fname = f"{self.prefix.get()}{num:03d}_{ver:02d}_mis{mis_ver:02d}.wav"
                    else:
                        count_dict[num] = count_dict.get(num, 0) + 1
                        ver = count_dict[num]
                        fname = f"{self.prefix.get()}{num:03d}_{ver:02d}.wav"
                    # 台詞欄の表示内容を分岐
                    if not m.get('is_mis'):
                        text_disp = scripts[[s['no'] for s in scripts].index(f"{num:03d}")]['text'].rstrip('。')
                    else:
                        text_disp = m['text'].rstrip('。')
                    no_disp = f"{num}(mis)" if m.get('is_mis') else num
                clip = audio[int(m['start']*1000): int(m['end']*1000)]
                clip.export(wav_sub / fname, format='wav')
                self.table.insert('', tk.END, values=(no_disp, text_disp, fname))

        messagebox.showinfo('完了', 'WAV出力が完了しました')
        # 進行中インジケーター解除
        self.config(cursor='')
        if hasattr(self, 'progress_label'):
            self.progress_label.config(text='')
        self.update()

    def play_selected(self):
        selected_items = self.table.selection()
        if not selected_items:
            messagebox.showinfo('情報', '再生する行を選択してください')
            return
            
        if self.is_playing:
            self.stop_playback()
            
        item = selected_items[0]
        values = self.table.item(item)['values']
        wav_name = values[2]
        
        # 出力ディレクトリから音声ファイルを探す
        wav_path = None
        for root, dirs, files in os.walk(self.out_dir.get()):
            if wav_name in files:
                wav_path = os.path.join(root, wav_name)
                break
                
        if wav_path:
            self.current_audio_path = wav_path
            self.play_thread = threading.Thread(target=self._play_audio)
            self.play_thread.start()
        else:
            messagebox.showerror('エラー', '音声ファイルが見つかりません')

    def play_continuous(self):
        selected_items = self.table.selection()
        if not selected_items:
            messagebox.showinfo('情報', '再生開始する行を選択してください')
            return
            
        if self.is_playing:
            self.stop_playback()
            
        self.continuous_play = True
        self.play_thread = threading.Thread(target=self._play_continuous)
        self.play_thread.start()

    def _play_continuous(self):
        selected_items = self.table.selection()
        current_index = self.table.index(selected_items[0])
        while self.continuous_play and current_index < len(self.table.get_children()):
            item = self.table.get_children()[current_index]
            values = self.table.item(item)['values']
            wav_name = values[2]
            wav_path = None
            for root, dirs, files in os.walk(self.out_dir.get()):
                if wav_name in files:
                    wav_path = os.path.join(root, wav_name)
                    break
            if wav_path:
                try:
                    print(f"連続再生: {wav_path}")
                    audio = AudioSegment.from_file(wav_path)
                    samples = np.array(audio.get_array_of_samples())
                    if audio.channels == 2:
                        samples = samples.reshape((-1, 2))
                    max_val = float(2 ** (8 * audio.sample_width - 1))
                    samples = samples.astype(np.float32) / max_val
                    sd.play(samples, audio.frame_rate)
                    sd.wait()
                except Exception as e:
                    print(f'連続再生エラー: {e}')
                    messagebox.showerror('エラー', f'音声の再生中にエラーが発生しました：\n{str(e)}')
            current_index += 1
            if current_index < len(self.table.get_children()):
                self.table.selection_set(self.table.get_children()[current_index])
                self.table.see(self.table.get_children()[current_index])

    def stop_playback(self):
        self.is_playing = False
        self.continuous_play = False
        sd.stop()

    def _play_audio(self):
        self.is_playing = True
        try:
            print(f"再生ファイル: {self.current_audio_path}")
            audio = AudioSegment.from_file(self.current_audio_path)
            samples = np.array(audio.get_array_of_samples())
            if audio.channels == 2:
                samples = samples.reshape((-1, 2))
            max_val = float(2 ** (8 * audio.sample_width - 1))
            samples = samples.astype(np.float32) / max_val
            sd.play(samples, audio.frame_rate)
            sd.wait()
        except Exception as e:
            print(f'再生エラー: {e}')
            messagebox.showerror('エラー', f'音声の再生中にエラーが発生しました：\n{str(e)}')
        finally:
            self.is_playing = False

    def on_select(self, event):
        if self.is_playing:
            self.stop_playback()

    def on_double_click(self, event):
        # クリックされた領域を取得
        region = self.table.identify_region(event.x, event.y)
        if region != "cell":
            return
            
        # クリックされた項目と列を取得
        column = self.table.identify_column(event.x)
        item = self.table.identify_row(event.y)
        
        if not item or not column:
            return
            
        # 列番号を取得
        column_num = int(column[1]) - 1
        
        # 編集可能な列かチェック
        if column_num not in [0, 1]:  # No列と台詞列のみ編集可能
            return
            
        # 既存の編集をキャンセル
        if self.edit_entry:
            self.edit_entry.destroy()
            
        # 項目の値を取得
        values = self.table.item(item)['values']
        current_value = values[column_num]
        
        # 編集用のエントリーを作成
        x, y, width, height = self.table.bbox(item, column)
        
        self.edit_entry = tk.Entry(self.table, width=width//10)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        
        self.edit_item = item
        self.edit_column = column_num
        
        # フォーカスを設定
        self.edit_entry.focus_set()
        
        # 編集完了時のイベントをバインド
        self.edit_entry.bind('<Return>', self.on_edit_complete)
        self.edit_entry.bind('<Escape>', self.on_edit_cancel)
        self.edit_entry.bind('<FocusOut>', self.on_edit_complete)

    def on_edit_complete(self, event):
        if not self.edit_entry:
            return
            
        # 新しい値を取得
        new_value = self.edit_entry.get()
        
        # 現在の値を取得
        values = list(self.table.item(self.edit_item)['values'])
        
        # 値を更新
        values[self.edit_column] = new_value
        self.table.item(self.edit_item, values=values)
        
        # 編集用エントリーを削除
        self.edit_entry.destroy()
        self.edit_entry = None
        self.edit_item = None
        self.edit_column = None

    def on_edit_cancel(self, event):
        if self.edit_entry:
            self.edit_entry.destroy()
            self.edit_entry = None
            self.edit_item = None
            self.edit_column = None

    def export_csv(self):
        if not self.table.get_children():
            messagebox.showinfo('情報', '出力するデータがありません')
            return
            
        # 保存先を選択
        file_path = filedialog.asksaveasfilename(
            defaultextension='.csv',
            filetypes=[('CSVファイル', '*.csv')],
            title='CSVファイルの保存'
        )
        
        if not file_path:
            return
            
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # ヘッダーを書き込み
                writer.writerow(['No', '台詞', '出力wav名'])
                
                # データを書き込み
                for item in self.table.get_children():
                    values = self.table.item(item)['values']
                    writer.writerow(values)
                    
            messagebox.showinfo('完了', 'CSVファイルの出力が完了しました')
        except Exception as e:
            messagebox.showerror('エラー', f'CSVファイルの出力中にエラーが発生しました：\n{str(e)}')

    def save_project(self):
        # 保存先を選択
        file_path = filedialog.asksaveasfilename(
            defaultextension='.json',
            filetypes=[('プロジェクトファイル', '*.json')],
            title='プロジェクトの保存'
        )
        
        if not file_path:
            return
            
        try:
            # プロジェクトデータの作成
            project_data = {
                'cv_path': self.cv_path.get(),
                'sheet': self.sheet.get(),
                'no_col': self.no_col.get(),
                'text_col': self.text_col.get(),
                'wav_dir': self.wav_dir.get(),
                'start_row': self.start_row.get(),
                'prefix': self.prefix.get(),
                'out_dir': self.out_dir.get(),
                'threshold': self.threshold.get(),
                'volume_thresh': self.volume_thresh.get(),
                'table_data': []
            }
            
            # テーブルデータの保存
            for item in self.table.get_children():
                values = self.table.item(item)['values']
                project_data['table_data'].append(values)
                
            # JSONファイルとして保存
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project_data, f, ensure_ascii=False, indent=2)
                
            messagebox.showinfo('完了', 'プロジェクトの保存が完了しました')
        except Exception as e:
            messagebox.showerror('エラー', f'プロジェクトの保存中にエラーが発生しました：\n{str(e)}')

    def load_project(self):
        # プロジェクトファイルを選択
        file_path = filedialog.askopenfilename(
            filetypes=[('プロジェクトファイル', '*.json')],
            title='プロジェクトを開く'
        )
        
        if not file_path:
            return
            
        try:
            # JSONファイルを読み込み
            with open(file_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
                
            # 設定を復元
            self.cv_path.set(project_data['cv_path'])
            self.sheet.set(project_data['sheet'])
            self.no_col.set(project_data['no_col'])
            self.text_col.set(project_data['text_col'])
            self.wav_dir.set(project_data['wav_dir'])
            self.start_row.set(project_data['start_row'])
            self.prefix.set(project_data['prefix'])
            self.out_dir.set(project_data['out_dir'])
            self.threshold.set(project_data['threshold'])
            self.volume_thresh.set(project_data['volume_thresh'])
            
            # テーブルデータを復元
            self.table.delete(*self.table.get_children())
            for values in project_data['table_data']:
                self.table.insert('', tk.END, values=values)
                
            messagebox.showinfo('完了', 'プロジェクトの読み込みが完了しました')
        except Exception as e:
            messagebox.showerror('エラー', f'プロジェクトの読み込み中にエラーが発生しました：\n{str(e)}')


if __name__ == '__main__':
    App().mainloop()
