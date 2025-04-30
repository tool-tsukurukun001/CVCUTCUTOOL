"""CVCUTCUTOOL – Tkinter GUI (出力フォルダ指定対応)
標準 Tkinter + 既存モジュールのみで動作
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from utils import load_script_list
from recognizer import transcribe_full
from matcher import match_segments
from naming import generate_filename
from pydub import AudioSegment
import threading

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('CVCUTCUTOOL')
        self.geometry('1050x650')

        # ------- 左ペイン入力 -------
        frm = tk.Frame(self)
        frm.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        # 0. CVリスト
        self.cv_path = tk.StringVar()
        tk.Label(frm, text='CVリスト').grid(row=0, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.cv_path, width=42).grid(row=0, column=1)
        tk.Button(frm, text='参照', command=self.browse_cv).grid(row=0, column=2)

        # 1. シート名
        self.sheet = tk.StringVar(value='Sheet1')
        tk.Label(frm, text='シート').grid(row=1, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.sheet, width=12).grid(row=1, column=1, sticky='w')

        # 2. 列指定
        self.no_col = tk.StringVar(value='A')
        self.text_col = tk.StringVar(value='B')
        tk.Label(frm, text='No列').grid(row=2, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.no_col, width=5).grid(row=2, column=1, sticky='w')
        tk.Label(frm, text='台詞列').grid(row=3, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.text_col, width=5).grid(row=3, column=1, sticky='w')

        # 3. WAV フォルダ
        self.wav_dir = tk.StringVar()
        tk.Label(frm, text='wavフォルダ').grid(row=4, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.wav_dir, width=42).grid(row=4, column=1)
        tk.Button(frm, text='参照', command=self.browse_wav).grid(row=4, column=2)

        # 4. 開始行
        self.start_row = tk.StringVar(value='2')
        tk.Label(frm, text='開始行').grid(row=5, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.start_row, width=8).grid(row=5, column=1, sticky='w')

        # 5. 出力ファイル名プレフィックス
        self.prefix = tk.StringVar(value='CVCUTCUTOOL_')
        tk.Label(frm, text='出力ファイル名プレフィックス').grid(row=6, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.prefix, width=25).grid(row=6, column=1, sticky='w')

        # 6. 出力先フォルダ
        self.out_dir = tk.StringVar()
        tk.Label(frm, text='出力先フォルダ').grid(row=7, column=0, sticky='e')
        tk.Entry(frm, textvariable=self.out_dir, width=42).grid(row=7, column=1)
        tk.Button(frm, text='参照', command=self.browse_out).grid(row=7, column=2)

        # 7. しきい値スライダー
        self.threshold = tk.IntVar(value=80); tk.Scale(frm, from_=0, to=100, orient='horizontal', variable=self.threshold,
                                                       label='しきい値 (ゆるい 0 〜 100 きびしい)', length=250).grid(row=8, column=0, columnspan=3, pady=5)

        # 8. 実行ボタン
        tk.Button(frm, text='WAV 出力', bg='royalblue', fg='white', width=22,
                  command=self.run).grid(row=9, column=0, columnspan=3, pady=15)

        # ------- 右ペイン テーブル -------
        cols = ('No', '台詞', '出力wav名')
        self.table = ttk.Treeview(self, columns=cols, show='headings', height=28)
        for c in cols:
            self.table.heading(c, text=c)
            self.table.column(c, width=260 if c=='台詞' else 140, anchor='w')
        self.table.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

    # ---------- ファイルダイアログ ----------
    def browse_cv(self):
        path = filedialog.askopenfilename(filetypes=[('Excel','*.xlsx')])
        if path:
            self.cv_path.set(path)

    def browse_wav(self):
        path = filedialog.askdirectory()
        if path:
            self.wav_dir.set(path)

    def browse_out(self):
        path = filedialog.askdirectory()
        if path:
            self.out_dir.set(path)

    # ---------- 実行 ----------
    def run(self):
        if not self.cv_path.get() or not self.wav_dir.get():
            messagebox.showerror('エラー','CVリストとwavフォルダを指定してください')
            return
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        self.table.delete(*self.table.get_children())
        cv_path = Path(self.cv_path.get())
        start_row = int(self.start_row.get() or 2)
        scripts = load_script_list(cv_path, sheet_name=self.sheet.get(), start_row=start_row)
        wav_dir = Path(self.wav_dir.get())
        out_root = Path(self.out_dir.get()) if self.out_dir.get() else Path.cwd()/ 'output'
        out_root.mkdir(exist_ok=True)

        for wav in wav_dir.glob('*.wav'):
            wav_sub = out_root / wav.stem
            wav_sub.mkdir(exist_ok=True)
            segments = transcribe_full(str(wav))
            matches  = match_segments(segments, scripts, threshold=80)
            audio    = AudioSegment.from_file(wav)
            for idx,m in enumerate(matches, start=1):
                start_ms = int(m['start']*1000)
                end_ms   = int(m['end']*1000)
                clip = audio[start_ms:end_ms]
                fname = f"{self.prefix.get()}{int(m['script_number']):03d}_{idx:02d}.wav"
                clip.export(wav_sub / fname, format='wav')
                self.table.insert('', tk.END, values=(m['script_number'], m['text'], fname))
        messagebox.showinfo('完了','WAV出力が完了しました')

if __name__ == '__main__':
    App().mainloop()