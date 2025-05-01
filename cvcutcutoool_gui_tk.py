import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from utils import load_script_list
from recognizer import transcribe_full
from matcher import match_segments
from pydub import AudioSegment
from fuzzywuzzy import fuzz
import threading


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CVCUTCUTOOL')
        self.geometry('1050x650')

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

        # ---------- 右ペイン：結果テーブル ----------
        cols = ('No', '台詞', '出力wav名')
        self.table = ttk.Treeview(self, columns=cols, show='headings', height=28)
        for c in cols:
            self.table.heading(c, text=c)
            self.table.column(c, width=260 if c == '台詞' else 140, anchor='w')
        self.table.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

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

            # 全体音声読み込み
            audio = AudioSegment.from_file(wav)

            # 1) Whisper で分割 & 認識
            segments = transcribe_full(str(wav))

            # 2) 音量フィルタ
            vol_th = self.volume_thresh.get()
            filtered = []
            for seg in segments:
                clip = audio[int(seg['start']*1000): int(seg['end']*1000)]
                if clip.dBFS >= vol_th:
                    filtered.append(seg)
            segments = filtered

            # 3) スクリプト照合
            matches = match_segments(segments, scripts, threshold=self.threshold.get())
            matched_keys = {(m['start'], m['end']) for m in matches}

            count_dict = {}
            # 正常マッチ
            for m in matches:
                num = int(m['script_number'])
                count_dict[num] = count_dict.get(num, 0) + 1
                ver = count_dict[num]
                fname = f"{self.prefix.get()}{num:03d}_{ver:02d}.wav"
                clip = audio[int(m['start']*1000): int(m['end']*1000)]
                clip.export(wav_sub / fname, format='wav')
                self.table.insert('', tk.END, values=(num, m['text'], fname))

            # mis 出力
            for seg in segments:
                key = (seg['start'], seg['end'])
                if key in matched_keys:
                    continue
                # best_script は文字列を返すので int に変換
                num_str = self.best_script(seg['text'], scripts)
                num = int(num_str)
                count_dict[num] = count_dict.get(num, 0) + 1
                ver = count_dict[num]
                fname = f"{self.prefix.get()}{num:03d}_{ver:02d}_mis.wav"
                clip = audio[int(seg['start']*1000): int(seg['end']*1000)]
                clip.export(wav_sub / fname, format='wav')
                # 表示は元の番号文字列＋"(mis)" に
                self.table.insert('', tk.END,
                                  values=(f"{num_str}(mis)", seg['text'], fname))

        messagebox.showinfo('完了', 'WAV出力が完了しました')


if __name__ == '__main__':
    App().mainloop()
