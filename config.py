from pathlib import Path

class Settings:
    # プロジェクトの基準パス
    BASE_DIR = Path(__file__).parent
    # 音声入力フォルダ
    AUDIO_DIR = BASE_DIR / "input_audio"
    # 台詞リスト Excel ファイル
    SCRIPT_EXCEL = BASE_DIR / "cv_list.xlsx"
    # 無音検出しきい値 (dBFS)
    SILENCE_THRESH = -40
    # 無音長さの最小値 (ms)
    MIN_SILENCE_LEN = 200

settings = Settings()
