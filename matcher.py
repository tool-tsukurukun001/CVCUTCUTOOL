from fuzzywuzzy import process, fuzz
import re
from unidecode import unidecode

def is_japanese_or_english(text):
    # ひらがな・カタカナ・英字・数字・記号のみを許可（漢字は除外）
    return bool(re.fullmatch(r'[ぁ-んァ-ンa-zA-Z0-9 　、。！？ー～「」『』]*', text))

def correct_text(text: str) -> str:
    return te


def match_segments(
    segments: list[dict],
    scripts: list[dict],
    threshold: int = 80
) -> list[dict]:
    matches = []
    script_texts = [s['text'] for s in scripts]
    n = len(segments)
    mis_count = 0
    for i in range(n):
        seg_text = segments[i]['text']
        best_match, score = process.extractOne(seg_text, script_texts, scorer=fuzz.ratio)
        if score >= threshold:
            script_no = scripts[script_texts.index(best_match)]['no']
            matches.append({
                'start': segments[i]['start'],
                'end': segments[i]['end'],
                'text': seg_text,
                'script_number': script_no,
                'score': score,
                'is_mis': False
            })
        else:
            mis_count += 1
            mis_text = unidecode(seg_text)
            matches.append({
                'start': segments[i]['start'],
                'end': segments[i]['end'],
                'text': mis_text,
                'script_number': '000',
                'score': score,
                'is_mis': True,
                'mis_index': mis_count
            })
    return matches
