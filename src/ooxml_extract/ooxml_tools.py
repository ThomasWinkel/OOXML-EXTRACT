from pathlib import Path
from oletools.olevba import VBA_Parser


def extract_vba_project(ooxml_path: Path, output_dir: Path) -> bool:
    vbaparser = VBA_Parser(ooxml_path)
    
    if not vbaparser.detect_vba_macros():
        return False
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
        module_file = output_dir / f"{vba_filename}"
        with open(module_file, 'w', encoding='utf-8', newline='',  errors='ignore') as f:
            f.write(vba_code)
            
    vbaparser.close()
    return True


def update_vba_project_bin(vba_project_bin: Path, vba_files: Path) -> bool:
    """
    Aktualisiert den Code in vbaProject.bin.
    TODO: Das scheint sehr schwierig zu sein. Das lasse ich erstmal weg.
    """
    pass
