def generate_filename(script_number: str, take_number: int) -> str:
    """
    script_number: 台詞リストの番号（文字列）
    take_number: 切り出し順の番号（1,2,3…）
    → 例: CVCUTCUTOOL_3_1.wav
    """
    return f"CVCUTCUTOOL_{script_number}_{take_number}.wav"
