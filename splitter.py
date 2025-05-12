from pydub import AudioSegment
from pydub.silence import detect_nonsilent
from typing import List, Dict
import os
import numpy as np
from scipy.signal import find_peaks


def detect_envelope(audio_segment: AudioSegment, threshold: float = 0.1) -> int:
    """
    音声のエンベロープを検出し、実際の発声開始位置を特定する
    
    Parameters:
    - audio_segment: 音声セグメント
    - threshold: エンベロープの閾値（0.0-1.0）
    
    Returns:
    - 発声開始位置（ミリ秒）
    """
    # 音声データをnumpy配列に変換
    samples = np.array(audio_segment.get_array_of_samples())
    if audio_segment.channels == 2:
        samples = samples.reshape((-1, 2))
        samples = np.mean(samples, axis=1)
    
    # エンベロープの計算
    envelope = np.abs(samples)
    envelope = np.convolve(envelope, np.ones(100)/100, mode='same')
    
    # エンベロープの正規化
    envelope = envelope / np.max(envelope)
    
    # 閾値を超える最初の位置を検出
    start_idx = np.where(envelope > threshold)[0]
    if len(start_idx) > 0:
        return int(start_idx[0] * 1000 / audio_segment.frame_rate)
    return 0


def apply_crossfade(audio_segment: AudioSegment, fade_duration: int = 10) -> AudioSegment:
    """
    音声セグメントにクロスフェードを適用する
    
    Parameters:
    - audio_segment: 音声セグメント
    - fade_duration: フェード時間（ミリ秒）
    
    Returns:
    - クロスフェード適用後の音声セグメント
    """
    return audio_segment.fade_in(fade_duration).fade_out(fade_duration)


def split_audio(
    input_path: str,
    min_silence_len: int = 500,
    silence_thresh: int = -50,
    keep_silence: int = 200,
    envelope_threshold: float = 0.1,
    crossfade_duration: int = 20
) -> List[Dict]:
    """
    音声ファイルを無音区間で分割し、音声チャンクのリストを返す。

    Parameters:
    - input_path: 入力WAVファイルのパス
    - min_silence_len: 無音とみなす最小の連続静寂長 (ms)
    - silence_thresh: 無音とみなす閾値 (dBFS) - 固定値 -50
    - keep_silence: 前後に保持する無音長 (ms)
    - envelope_threshold: エンベロープ検出の閾値
    - crossfade_duration: クロスフェード時間 (ms) - 固定値 20

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
        
        # エンベロープ検出による開始位置の調整
        envelope_start = detect_envelope(chunk, envelope_threshold)
        if envelope_start > 0:
            chunk = chunk[envelope_start:]
            seg_start += envelope_start
        
        # クロスフェードの適用
        chunk = apply_crossfade(chunk, crossfade_duration)
        
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
