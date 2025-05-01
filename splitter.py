from pydub import AudioSegment
from pydub.silence import detect_nonsilent
from typing import List, Dict
import os


def split_audio(
    input_path: str,
    min_silence_len: int = 500,
    silence_thresh: int = -40,
    keep_silence: int = 200
) -> List[Dict]:
    """
    音声ファイルを無音区間で分割し、音声チャンクのリストを返す。

    Parameters:
    - input_path: 入力WAVファイルのパス
    - min_silence_len: 無音とみなす最小の連続静寂長 (ms)
    - silence_thresh: 無音とみなす閾値 (dBFS)
    - keep_silence: 前後に保持する無音長 (ms)

    Returns:
    List of dicts:
      {
        'audio': AudioSegmentチャンク,
        'start': 開始時刻(秒),
        'end'  : 終了時刻(秒)
      }
    """
    audio = AudioSegment.from_file(input_path)
    # 無音でない区間の検出 (ms単位)
    nonsilent_ranges = detect_nonsilent(
        audio,
        min_silence_len=min_silence_len,
        silence_thresh=silence_thresh
    )

    segments: List[Dict] = []
    for start_i, end_i in nonsilent_ranges:
        # 前後に silence を付加して切り出し
        seg_start = max(0, start_i - keep_silence)
        seg_end = min(len(audio), end_i + keep_silence)
        chunk = audio[seg_start:seg_end]
        segments.append({
            'audio': chunk,
            'start': seg_start / 1000.0,
            'end': seg_end / 1000.0
        })
    return segments


def save_segments(
    segments: List[Dict],
    output_dir: str,
    prefix: str = 'seg'
) -> None:
    """
    分割したセグメントを WAV ファイルとして保存する。

    - output_dir: 保存先ディレクトリ
    - prefix: ファイル名プレフィックス
    """
    os.makedirs(output_dir, exist_ok=True)
    for idx, seg in enumerate(segments, start=1):
        fname = f"{prefix}_{idx:03d}.wav"
        out_path = os.path.join(output_dir, fname)
        seg['audio'].export(out_path, format='wav')
