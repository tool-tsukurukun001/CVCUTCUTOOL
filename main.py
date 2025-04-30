import argparse
from pathlib import Path
from config import settings
from utils import load_script_list
from recognizer import transcribe_full, transcribe_chunks
from matcher import match_segments
from naming import generate_filename
from pydub import AudioSegment

def main():
    parser = argparse.ArgumentParser(description="CVCUTCUTOOL: CV Auto Split & Rename Tool")
    parser.add_argument("--audio-dir", type=Path, default=settings.AUDIO_DIR)
    parser.add_argument("--excel", type=Path, default=settings.SCRIPT_EXCEL)
    parser.add_argument("--out-dir", type=Path, default=settings.BASE_DIR / 'output')
    parser.add_argument("--model", type=str, default="base")
    parser.add_argument("--threshold", type=int, default=10, help="照合のしきい値（デフォルト: 10）")
    args = parser.parse_args()

    # 出力フォルダ作成
    args.out_dir.mkdir(exist_ok=True)

    # 台詞リスト読み込み
    scripts = load_script_list(args.excel)

    for wav in args.audio_dir.glob("*.wav"):
        print(f"[Debug] Processing {wav.name} ...")

        # WAVごとにサブフォルダを準備（既存ファイルはクリア）
        wav_dir = args.out_dir / wav.stem
        if wav_dir.exists():
            for f in wav_dir.iterdir():
                f.unlink()
        else:
            wav_dir.mkdir(parents=True)

        # 全体文字起こし
        segments = transcribe_full(str(wav), model_size=args.model)
        for seg in segments:
            print(f"[Full] {seg['start']:.2f}s → {seg['end']:.2f}s: {seg['text']}")

        # 台詞リストと照合して切り出し
        matches = match_segments(segments, scripts, threshold=args.threshold)
        audio = AudioSegment.from_file(wav)
        for idx, m in enumerate(matches, start=1):
            start_ms = int(m["start"] * 1000)
            end_ms   = int(m["end"]   * 1000)
            clip = audio[start_ms:end_ms]
            out_name = generate_filename(m["script_number"], idx)
            clip.export(wav_dir / out_name, format="wav")
            print(f"[Cut] {out_name} ({m['text']}, {m['start']:.2f}-{m['end']:.2f}s)")

        # （オプション）分割チャンク→再文字起こしで確認  
        # chunks = split_audio(wav, wav_dir, settings.SILENCE_THRESH, settings.MIN_SILENCE_LEN)  
        # transcriptions = transcribe_chunks(wav_dir)  
        # for fname, text in transcriptions.items():  
        #     print(f"[Result] {fname} → {text}")

if __name__ == "__main__":
    main()
