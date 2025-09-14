import json
import os
from pathlib import Path

APP_NAME = "DocuDrop"
DEFAULT_SETTINGS = {
    "captions_enabled": True,
    "caption_prefix": "Figure",
    "caption_custom_text": "",
    "numbering_format": "decimal",
    "caption_style": {"italic": True, "size_pt": 9},
    "caption_min_words": 3,
    "caption_max_words": 7,
    "tof_enabled": True,
    "tof_position": "start",

    # NEW
    "use_codeword_insertion": False,
    "insertion_codeword": "[[IMG]]",

    # Screenshot Copy
    "save_copy_enabled": False,
    "save_copy_folder": "",

    # Image Compression
    "compression_enabled": False,
    "compression_quality": 85
}

def get_settings_path():
    # Prefer per-user local appdata on Windows, otherwise home dir
    if os.name == "nt":
        base = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
    else:
        base = os.path.expanduser("~")
    config_dir = Path(base) / APP_NAME
    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir / "settings.json"

def load_settings():
    path = get_settings_path()
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            # merge defaults with loaded so new fields get default values
            merged = DEFAULT_SETTINGS.copy()
            merged.update(data)
            return merged
        except Exception:
            return DEFAULT_SETTINGS.copy()
    else:
        return DEFAULT_SETTINGS.copy()

def save_settings(settings: dict):
    path = get_settings_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
        return True
    except Exception:
        return False

