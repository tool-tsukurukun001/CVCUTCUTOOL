from fuzzywuzzy import process

def match_segments(segments, script_list, threshold=80):
    """
    segments: [
        {"start": float, "end": float, "text": str}, ...
    ]
    script_list: {'1':'おはようございます。', '2':'よろしくお願いします。', ...}
    threshold: 照合の一致度しきい値
    → 戻り値: [
        {
            "script_number": key,
            "start": float,
            "end": float,
            "text": str,
            "score": int
        },
        ...
    ]
    """
    matches = []
    for i, seg in enumerate(segments):
        text = seg["text"]

        # 1) 優先キーワードマッチ例（必要なら追加）
        # if "オーバーフローエラー" in text:
        #     matches.append({...}); continue

        # 2) 隣接セグメントを結合してより長い台詞にマッチするか試す
        if i + 1 < len(segments):
            combined = text + segments[i + 1]["text"]
            best_comb, comb_score = process.extractOne(combined, list(script_list.values()))
            if comb_score >= threshold and len(best_comb) > len(text):
                key = next(k for k, v in script_list.items() if v == best_comb)
                matches.append({
                    "script_number": key,
                    "start": seg["start"],
                    "end": segments[i + 1]["end"],
                    "text": best_comb,
                    "score": comb_score
                })
                # 次のセグメントはスキップ
                continue

        # 3) 通常のあいまい照合
        best, score = process.extractOne(text, list(script_list.values()))
        if score >= threshold:
            key = next(k for k, v in script_list.items() if v == best)
            matches.append({
                "script_number": key,
                "start": seg["start"],
                "end": seg["end"],
                "text": best,
                "score": score
            })
    return matches
