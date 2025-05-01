from fuzzywuzzy import process


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
          'score': int
        },
        ...
      ]
    """
    matches = []
    # 台本テキスト一覧を作成
    script_texts = [s['text'] for s in scripts]

    for seg in segments:
        seg_text = seg.get('text', '')
        # 最も類似する台本テキストを取得
        best_text, score = process.extractOne(seg_text, script_texts)
        if score >= threshold:
            # マッチしたテキストに対応する台詞番号を取得
            idx = script_texts.index(best_text)
            script_no = scripts[idx]['no']
            matches.append({
                'start': seg['start'],
                'end': seg['end'],
                'text': seg_text,
                'script_number': script_no,
                'score': score
            })
    return matches
