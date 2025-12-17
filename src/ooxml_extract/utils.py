from pathlib import Path


def get_unique_folder_name(base_path: Path) -> Path:
    """
    Gibt einen eindeutigen Ordnernamen zurück.
    Wenn der Ordner existiert, wird _001, _002, etc. angehängt.
    """
    if not base_path.exists():
        return base_path
    
    counter = 1
    while True:
        new_path = base_path.parent / f"{base_path.name}_{counter:03d}"
        if not new_path.exists():
            return new_path
        counter += 1
