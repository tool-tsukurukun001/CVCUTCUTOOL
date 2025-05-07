from fuzzywuzzy import process
from fuzzywuzzy import fuzz


def match_segments(
    segments: list[dict],
    scripts: list[dict],
    threshold: int = 80
) -> list[dict]:
    """
    Whisper の出力セグメントと台本リストを照合し、閾値以上のマッチを返す。

    Parameters:
    - segments: [
        {'start': float, 'end': float, 'text': str}, ...
      ]
    - scripts: [
        {'no': '001', 'text': 'セリフ本文'}, ...
      ]
    - threshold: 0-100 の類似度しきい値

    Returns:
    - matches: [
        {
          'start': float,
          'end': float,
          'text': str,
          'script_number': '001',
          'score': int,
          'is_mis': bool
        },
        ...
      ]
    """
    matches = []
    script_texts = [s['text'] for s in scripts]

    for seg in segments:  # 元の順番で処理
        seg_text = seg.get('text', '')
        if not seg_text:
            continue
        best_text, score = process.extractOne(seg_text, script_texts)
        idx = script_texts.index(best_text)
        script_no = scripts[idx]['no']
        is_mis = score < threshold
        matches.append({
            'start': seg['start'],
            'end': seg['end'],
            'text': seg_text,
            'script_number': script_no,
            'score': score,
            'is_mis': is_mis
        })
    return matches
