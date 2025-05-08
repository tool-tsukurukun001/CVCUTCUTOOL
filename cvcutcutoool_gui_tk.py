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
from PIL import Image, ImageTk
import queue


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CVCUTCUTOOL')
        self.geometry('1050x650')

        # 再生用の変数
        self.current_audio = None
        self.is_playing = False
        self.play_thread = None

        # 背景テクスチャ
        try:
            self.bg_img = Image.open('texture.png')
            self.bg_img = self.bg_img.resize((1050, 650))
            self.bg_photo = ImageTk.PhotoImage(self.bg_img)
            self.bg_label = tk.Label(self, image=self.bg_photo)
            self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            self.bg_label.lower()
        except Exception as e:
            print(f'背景テクスチャ読み込み失敗: {e}')

        # テクスチャに近い色
        tex_color = '#f5f5f5'

        # ---------- 左ペイン ----------
        frm = tk.Frame(self, bg=tex_color)
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
        self.threshold = tk.IntVar(value=50)  # デフォルト50
        tk.Scale(
            frm, from_=0, to=100, orient='horizontal', variable=self.threshold,
            label='しきい値 (ゆるい 0〜100 きびしい)', length=250
        ).grid(row=8, column=0, columnspan=3, pady=5)

        # 音量しきい値 (dBFS)
        self.volume_thresh = tk.IntVar(value=-50)  # デフォルト-50
        tk.Scale(
            frm, from_=-60, to=0, orient='horizontal', variable=self.volume_thresh,
            label='音量しきい値 (dBFS)', length=250
        ).grid(row=9, column=0, columnspan=3, pady=5)

        # プレビュー（旧WAV出力）ボタン
        tk.Button(
            frm, text='プレビュー', bg='royalblue', fg='white', width=22,
            command=self.preview
        ).grid(row=10, column=0, columnspan=3, pady=15)

        # WAV出力ボタン（プレビューの下に追加）
        self.export_wav_button = tk.Button(frm, text='WAV出力', bg='green', fg='white', width=22, command=self.export_wav_files)
        self.export_wav_button.grid(row=15, column=0, columnspan=3, pady=10)

        # --- ここから下にコントロールボタンを縦に配置 ---
        self.play_button = tk.Button(frm, text='再生 (S)', command=self.play_selected, width=22)
        self.play_button.grid(row=11, column=0, columnspan=3, pady=2)
        self.continuous_play_button = tk.Button(frm, text='連続再生 (R)', command=self.play_continuous, width=22)
        self.continuous_play_button.grid(row=12, column=0, columnspan=3, pady=2)
        self.stop_button = tk.Button(frm, text='停止 (0)', command=self.stop_playback, width=22)
        self.stop_button.grid(row=13, column=0, columnspan=3, pady=2)
        self.exclude_button = tk.Button(frm, text='除外 (X)', width=22, command=self.exclude_selected)
        self.exclude_button.grid(row=14, column=0, columnspan=3, pady=2)
        # プレビューボタン（青）
        self.preview_button = tk.Button(frm, text='プレビュー', bg='royalblue', fg='white', width=22, command=self.preview)
        self.preview_button.grid(row=15, column=0, columnspan=3, pady=10)
        # オールクリアボタン（赤）
        self.clear_button = tk.Button(frm, text='オールクリア', width=22, command=self.all_clear, bg='red', fg='white')
        self.clear_button.grid(row=16, column=0, columnspan=3, pady=2)
        # CSV/WAV出力ボタンを最下段に左右並びで配置
        bottom_frame = tk.Frame(frm, bg=tex_color)
        bottom_frame.grid(row=17, column=0, columnspan=3, pady=10, sticky='ew')
        self.csv_button = tk.Button(bottom_frame, text='CSV出力', command=self.export_csv, width=10)
        self.csv_button.pack(side=tk.LEFT, padx=5)
        self.export_wav_button = tk.Button(bottom_frame, text='WAV出力', command=self.export_wav_files, width=10)
        self.export_wav_button.pack(side=tk.RIGHT, padx=5)

        # ---------- 右ペイン：結果テーブル ----------
        cols = ('No', '台詞', '出力wav名', '開始', '終了', '元WAV')
        table_frame = tk.Frame(self, bg=tex_color)
        table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)
        style = ttk.Style()
        style.configure('Treeview', background=tex_color, fieldbackground=tex_color)
        self.table = ttk.Treeview(table_frame, columns=cols, show='headings', height=28, style='Treeview')
        for c in cols:
            self.table.heading(c, text=c)
            self.table.column(c, width=180 if c == '台詞' else 100, anchor='w')
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=vsb.set)
        self.table.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)

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
        file_menu.add_command(label='OKテイクにマーキング', command=self.mark_ok_take)

        # 除外フラグ用の辞書
        self.excluded_items = set()

        # ショートカットキーのバインド
        self.bind('s', self.on_s_key)
        self.bind('r', self.on_r_key)
        self.bind('0', self.on_0_key)
        self.bind('x', self.on_x_key)
        self.bind('k', self.on_k_key)

        # プレビュー用キューとスレッド管理
        self.preview_queue = queue.Queue()
        self.preview_thread = None
        self.after(100, self.process_preview_queue)

        # 各入力欄・ボタンの参照を保存
        self.cv_path_entry = frm.children[list(frm.children)[1]]
        self.cv_path_btn = frm.children[list(frm.children)[2]]
        self.sheet_entry = frm.children[list(frm.children)[4]]
        self.no_col_entry = frm.children[list(frm.children)[6]]
        self.text_col_entry = frm.children[list(frm.children)[8]]
        self.wav_dir_entry = frm.children[list(frm.children)[10]]
        self.wav_dir_btn = frm.children[list(frm.children)[11]]
        self.start_row_entry = frm.children[list(frm.children)[13]]
        self.prefix_entry = frm.children[list(frm.children)[15]]
        self.out_dir_entry = frm.children[list(frm.children)[17]]
        self.out_dir_btn = frm.children[list(frm.children)[18]]
        self.threshold_scale = frm.children[list(frm.children)[19]]
        self.volume_thresh_scale = frm.children[list(frm.children)[20]]

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

    # プレビュートリガー
    def preview(self):
        # プレビュー後は左ペインの入力欄・参照ボタンを無効化
        self.set_left_controls_state('disabled')
        if not self.cv_path.get() or not self.wav_dir.get():
            messagebox.showerror('エラー', 'CVリストとwavフォルダを指定してください')
            return
        self.config(cursor='watch')
        if not hasattr(self, 'progress_label'):
            self.progress_label = tk.Label(self, text='処理中...', fg='red', font=('Arial', 20), bg='#fff')
            self.progress_label.place(relx=0.5, rely=0.5, anchor='center')
        else:
            self.progress_label.config(text='処理中...')
            self.progress_label.place(relx=0.5, rely=0.5, anchor='center')
        self.update()
        self.table.delete(*self.table.get_children())
        # スレッドで1ファイルずつ処理
        self.preview_thread = threading.Thread(target=self.process_preview_async, daemon=True)
        self.preview_thread.start()

    def process_preview_async(self):
        scripts = load_script_list(
            Path(self.cv_path.get()),
            sheet_name=self.sheet.get(),
            start_row=int(self.start_row.get() or 2),
            no_col=self.no_col.get(),
            text_col=self.text_col.get()
        )
        wav_dir = Path(self.wav_dir.get())
        for wav in wav_dir.glob('*.wav'):
            wav_name = wav.name
            chunk_segments = split_audio(
                str(wav),
                min_silence_len=700,
                silence_thresh=-40,
                keep_silence=400
            )
            segments = []
            for idx, chunk in enumerate(chunk_segments, start=1):
                tmp_chunk_path = wav.parent / f"_tmp_chunk_{idx:03d}.wav"
                chunk['audio'].export(tmp_chunk_path, format='wav')
                chunk_result = transcribe_full(str(tmp_chunk_path))
                for seg in chunk_result:
                    seg_abs = {
                        'start': chunk['start'] + seg['start'],
                        'end': chunk['start'] + seg['end'],
                        'text': seg['text']
                    }
                    segments.append(seg_abs)
                tmp_chunk_path.unlink()
            vol_th = self.volume_thresh.get()
            filtered = []
            audio = AudioSegment.from_file(wav)
            for seg in segments:
                clip = audio[int(seg['start']*1000): int(seg['end']*1000)]
                if clip.dBFS >= vol_th:
                    filtered.append(seg)
            segments = filtered
            matches = match_segments(segments, scripts, threshold=self.threshold.get())
            count_dict = {}
            mis_count_dict = {}
            for m in matches:
                if m.get('is_mis') and m['script_number'] == '000':
                    fname = f"{self.prefix.get()}000_00_mis{m['mis_index']:02d}.wav"
                    text_disp = m['text']
                    no_disp = '000(mis)'
                else:
                    num = int(m['script_number'])
                    if m.get('is_mis'):
                        mis_count_dict[num] = mis_count_dict.get(num, 0) + 1
                        ver = count_dict.get(num, 0) + 1
                        mis_ver = mis_count_dict[num]
                        fname = f"{self.prefix.get()}{num}_{ver}_mis{mis_ver}.wav"
                    else:
                        count_dict[num] = count_dict.get(num, 0) + 1
                        ver = count_dict[num]
                        fname = f"{self.prefix.get()}{num}_{ver}.wav"
                    # No比較はstr(int(...))で統一
                    script_no_list = []
                    for s in scripts:
                        try:
                            script_no_list.append(str(int(s['no'])))
                        except Exception:
                            script_no_list.append(str(s['no']))
                    num_str = str(int(num)) if str(num).isdigit() else str(num)
                    if num_str in script_no_list:
                        text_disp = scripts[script_no_list.index(num_str)]['text'].rstrip('。')
                    else:
                        print(f'[WARN] No {num_str} not found in script_no_list: {script_no_list}')
                        text_disp = m['text'].rstrip('。')
                    no_disp = f"{num}(mis)" if m.get('is_mis') else num
                self.preview_queue.put((no_disp, text_disp, fname, m['start'], m['end'], wav_name))
        self.preview_queue.put('DONE')

    def process_preview_queue(self):
        try:
            while not self.preview_queue.empty():
                item = self.preview_queue.get_nowait()
                if item == 'DONE':
                    self.config(cursor='')
                    if hasattr(self, 'progress_label'):
                        self.progress_label.config(text='')
                        self.progress_label.place_forget()
                    self.update()
                    return
                self.table.insert('', tk.END, values=item)
        except Exception as e:
            print(f'プレビューキュー処理エラー: {e}')
        finally:
            self.after(100, self.process_preview_queue)

    def play_selected(self):
        selected_items = self.table.selection()
        if not selected_items:
            messagebox.showinfo('情報', '再生する行を選択してください')
            return
        if self.is_playing:
            self.stop_playback()
        item = selected_items[0]
        values = self.table.item(item)['values']
        fname = values[2]
        start = float(values[3])
        end = float(values[4])
        wav_name = values[5]
        wav_dir = Path(self.wav_dir.get())
        wav_path = wav_dir / wav_name
        if not wav_path.exists():
            messagebox.showerror('エラー', f'元WAVファイルが見つかりません: {wav_name}')
            return
        self.current_audio_path = str(wav_path)
        self.preview_start = start
        self.preview_end = end
        self.play_thread = threading.Thread(target=self._play_preview_audio)
        self.play_thread.start()

    def _play_preview_audio(self):
        self.is_playing = True
        try:
            print(f"プレビュー再生: {self.current_audio_path} {self.preview_start}-{self.preview_end}")
            audio = AudioSegment.from_file(self.current_audio_path)
            preview = audio[int(self.preview_start*1000):int(self.preview_end*1000)]
            samples = np.array(preview.get_array_of_samples())
            if preview.channels == 2:
                samples = samples.reshape((-1, 2))
            max_val = float(2 ** (8 * preview.sample_width - 1))
            samples = samples.astype(np.float32) / max_val
            sd.play(samples, preview.frame_rate)
            sd.wait()
        except Exception as e:
            print(f'プレビュー再生エラー: {e}')
            messagebox.showerror('エラー', f'プレビュー再生中にエラーが発生しました：\n{str(e)}')
        finally:
            self.is_playing = False

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
            fname = values[2]
            start = float(values[3])
            end = float(values[4])
            wav_name = values[5]
            wav_dir = Path(self.wav_dir.get())
            wav_path = wav_dir / wav_name
            if wav_path.exists():
                try:
                    print(f"連続プレビュー再生: {wav_path} {start}-{end}")
                    audio = AudioSegment.from_file(str(wav_path))
                    preview = audio[int(start*1000):int(end*1000)]
                    samples = np.array(preview.get_array_of_samples())
                    if preview.channels == 2:
                        samples = samples.reshape((-1, 2))
                    max_val = float(2 ** (8 * preview.sample_width - 1))
                    samples = samples.astype(np.float32) / max_val
                    sd.play(samples, preview.frame_rate)
                    sd.wait()
                except Exception as e:
                    print(f'連続プレビュー再生エラー: {e}')
                    messagebox.showerror('エラー', f'連続プレビュー再生中にエラーが発生しました：\n{str(e)}')
            current_index += 1
            if current_index < len(self.table.get_children()):
                self.table.selection_set(self.table.get_children()[current_index])
                self.table.see(self.table.get_children()[current_index])

    def stop_playback(self):
        self.is_playing = False
        self.continuous_play = False
        sd.stop()

    def on_select(self, event):
        if self.is_playing:
            self.stop_playback()

    def on_double_click(self, event):
        region = self.table.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.table.identify_column(event.x)
        item = self.table.identify_row(event.y)
        if not item or not column:
            return
        column_num = int(column[1]) - 1
        # 編集可能な列: No, 台詞, 出力wav名, 開始, 終了
        if column_num not in [0, 1, 2, 3, 4]:
            return
        if self.edit_entry:
            self.edit_entry.destroy()
        values = self.table.item(item)['values']
        current_value = values[column_num]
        x, y, width, height = self.table.bbox(item, column)
        self.edit_entry = tk.Entry(self.table, width=width//10)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.select_range(0, tk.END)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_item = item
        self.edit_column = column_num
        self.edit_entry.focus_set()
        self.edit_entry.bind('<Return>', self.on_edit_complete)
        self.edit_entry.bind('<Escape>', self.on_edit_cancel)
        self.edit_entry.bind('<FocusOut>', self.on_edit_complete)

    def on_edit_complete(self, event):
        if not self.edit_entry:
            return
        new_value = self.edit_entry.get()
        values = list(self.table.item(self.edit_item)['values'])
        # No列を編集した場合は出力wav名も自動更新
        if self.edit_column == 0:
            old_no = str(values[0])
            new_no = str(int(new_value)) if str(new_value).isdigit() else new_value
            # 既存Noグループのバージョン最大値を計算
            all_items = self.table.get_children()
            ver = 1
            for item in all_items:
                v = self.table.item(item)['values']
                if str(v[0]) == new_no and item != self.edit_item:
                    fname = v[2]
                    try:
                        ver_num = int(fname.split('_')[-1].split('.')[0])
                        if ver_num >= ver:
                            ver = ver_num + 1
                    except Exception:
                        continue
            # ファイル名を新No＋新バージョンで生成
            old_fname = values[2]
            parts = old_fname.split('_')
            if len(parts) >= 3:
                parts[1] = new_no
                parts[-1] = f'{ver}.wav'
                values[2] = '_'.join(parts)
            values[self.edit_column] = new_no
        # 台詞欄を編集した場合は最も近いNoに自動変換し、ファイル名も自動更新
        elif self.edit_column == 1:
            scripts = load_script_list(
                Path(self.cv_path.get()),
                sheet_name=self.sheet.get(),
                start_row=int(self.start_row.get() or 2),
                no_col=self.no_col.get(),
                text_col=self.text_col.get()
            )
            best_score = -1
            best_no = None
            for s in scripts:
                score = fuzz.ratio(new_value, s['text'])
                if score > best_score:
                    best_score = score
                    best_no = s['no']
            if best_no is not None:
                # mis解除時はファイル名からmisを除去し、Noグループの最大値＋1を割り当て
                new_no = str(int(best_no)) if str(best_no).isdigit() else best_no
                all_items = self.table.get_children()
                ver = 1
                for item in all_items:
                    v = self.table.item(item)['values']
                    if str(v[0]) == new_no and item != self.edit_item:
                        fname = v[2]
                        try:
                            ver_num = int(fname.split('_')[-1].split('.')[0])
                            if ver_num >= ver:
                                ver = ver_num + 1
                        except Exception:
                            continue
                old_fname = values[2]
                parts = old_fname.split('_')
                # mis解除時はmisを除去
                if 'mis' in parts[-1]:
                    parts = parts[:-1] + [f'{ver}.wav']
                else:
                    parts[1] = new_no
                    parts[-1] = f'{ver}.wav'
                values[2] = '_'.join(parts)
                values[0] = new_no
        values[self.edit_column] = new_value
        self.table.item(self.edit_item, values=values)
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
                writer.writerow(['No', '台詞', '出力wav名', '開始', '終了', '元WAV', 'OKテイク'])
                # データを書き込み
                for item in self.table.get_children():
                    values = self.table.item(item)['values']
                    tags = self.table.item(item, 'tags')
                    ok_flag = 'OK' if 'oktake' in tags else ''
                    writer.writerow(list(values) + [ok_flag])
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

    def export_wav_files(self):
        if not self.table.get_children():
            messagebox.showinfo('情報', '出力するデータがありません')
            return
        out_dir = self.out_dir.get() or (str(Path.cwd() / 'output'))
        Path(out_dir).mkdir(exist_ok=True)
        wav_dir = Path(self.wav_dir.get())
        for item in self.table.get_children():
            if item in self.excluded_items:
                continue  # 除外行はスキップ
            values = self.table.item(item)['values']
            fname = values[2]
            start = float(values[3])
            end = float(values[4])
            wav_name = values[5]
            wav_path = wav_dir / wav_name
            if not wav_path.exists():
                print(f'元WAVファイルが見つかりません: {wav_name}')
                continue
            audio = AudioSegment.from_file(str(wav_path))
            clip = audio[int(start*1000):int(end*1000)]
            out_path = Path(out_dir) / fname
            clip.export(out_path, format='wav')
            print(f'書き出し: {out_path}')
        messagebox.showinfo('完了', 'WAVファイルの出力が完了しました')

    def exclude_selected(self):
        selected_items = self.table.selection()
        for item in selected_items:
            if item not in self.excluded_items:
                self.table.item(item, tags=('excluded',))
                self.excluded_items.add(item)
            else:
                self.table.item(item, tags=())
                self.excluded_items.remove(item)
        self.table.tag_configure('excluded', background='#cccccc')

    def on_s_key(self, event):
        self.play_selected()

    def on_r_key(self, event):
        self.play_continuous()

    def on_0_key(self, event):
        self.stop_playback()

    def on_x_key(self, event):
        self.exclude_selected()

    def on_k_key(self, event):
        self.mark_ok_take()

    def mark_ok_take(self):
        selNoected_items = self.table.selection()
        for item in selected_items:
            tags = self.table.item(item, 'tags')
            if isinstance(tags, str):
                tags = (tags,)
            if 'oktake' in tags:
                self.table.item(item, tags=tuple(t for t in tags if t != 'oktake'))
            else:
                self.table.item(item, tags=tuple(tags) + ('oktake',))
        self.table.tag_configure('oktake', background='#fff2b2')  # 薄い黄色

    def set_left_controls_state(self, state):
        # CVリスト～出力先フォルダまでの入力欄・参照ボタンを有効/無効化
        for widget in [self.cv_path_entry, self.cv_path_btn, self.sheet_entry, self.no_col_entry, self.text_col_entry,
                       self.wav_dir_entry, self.wav_dir_btn, self.start_row_entry, self.prefix_entry, self.out_dir_entry, self.out_dir_btn]:
            widget.config(state=state)
        self.threshold_scale.config(state=state)
        self.volume_thresh_scale.config(state=state)

    def all_clear(self):
        if not messagebox.askyesno('確認', '本当に全消去してよろしいですか？'):
            return
        self.table.delete(*self.table.get_children())
        self.set_left_controls_state('normal')


if __name__ == '__main__':
    App().mainloop()
