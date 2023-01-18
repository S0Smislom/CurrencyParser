from pathlib import Path

def get_or_create_path(path, subpath):
    """Возвращает или создает путь"""
    filepath = Path(path / subpath)
    filepath.mkdir(parents=True, exist_ok=True)
    return filepath
