from pydub import AudioSegment, silence

def split_audio(input_path, output_dir, silence_thresh, min_silence_len):
    """
    input_path: Path またはファイルパス文字列
    output_dir: Path オブジェクトの出力先フォルダ
    silence_thresh: 無音とみなす dBFS（例: -20）
    min_silence_len: 無音判定の最小長さ（ミリ秒）
    """
    # 音声読み込み
    audio = AudioSegment.from_file(input_path)
    # 音量情報を出力
    print(f"[Debug] Loaded {input_path.name}, average dBFS: {audio.dBFS:.2f}")

    # 無音区間で分割
    chunks = silence.split_on_silence(
        audio,
        min_silence_len=min_silence_len,
        silence_thresh=silence_thresh
    )
    if not chunks:
        print(f"[Debug] No chunks detected (threshold={silence_thresh}, min_len={min_silence_len}ms)")

    # 分割チャンクをファイル出力
    for i, chunk in enumerate(chunks, start=1):
        out_file = output_dir / f"chunk_{i}.wav"
        chunk.export(out_file, format="wav")

    return chunks
