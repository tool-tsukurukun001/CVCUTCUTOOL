from faster_whisper import WhisperModel
from pathlib import Path

def transcribe_chunks(chunks_dir: Path, model_size: str = "base"):
    """
    chunks_dir: 分割チャンクが入ったフォルダ (Path)
    model_size: "tiny", "base", "small" など
    → 戻り値: { "chunk_1.wav": "テキスト...", … }
    """
    model = WhisperModel(model_size, device="cpu", compute_type="int8")
    results = {}
    for audio_file in sorted(chunks_dir.glob("chunk_*.wav")):
        print(f"[Transcribe] {audio_file.name} …")
        segments, _ = model.transcribe(str(audio_file))
        text = " ".join([segment.text for segment in segments])
        results[audio_file.name] = text.strip()
    return results

def transcribe_full(wav_path: str, model_size: str = "base"):
    """
    wav_path: 元の WAV ファイルパス
    model_size: "tiny", "base", "small" など
    → 戻り値: [
         {"start": float, "end": float, "text": str},
         …
       ]
    """
    model = WhisperModel(model_size, device="cpu", compute_type="int8")
    segments, _ = model.transcribe(wav_path)
    return [
        {"start": segment.start, "end": segment.end, "text": segment.text.strip()}
        for segment in segments
    ]
