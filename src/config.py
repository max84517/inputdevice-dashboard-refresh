"""Persistent configuration management using JSON."""
import json
import os
import sys

# When running as a PyInstaller frozen exe, __file__ points to the temporary
# extraction folder (_MEIxxxxxx) which is deleted on exit.  Use the directory
# that contains the exe instead so config.json persists next to the exe.
if getattr(sys, "frozen", False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.dirname(__file__))

CONFIG_PATH = os.path.join(_BASE_DIR, "config.json")


def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(data: dict):
    existing = load_config()
    existing.update(data)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(existing, f, indent=2, ensure_ascii=False)
