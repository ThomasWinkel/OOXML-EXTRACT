from pathlib import Path
import click
import zipfile
import shutil
from typing import Set
from .xml_formatter import prettify_xml_file, minify_xml_file_to_bytes
from .ooxml_vba import export_vba_project, import_vba_project
from .utils import get_unique_folder_name


def prettify_xml_files(directory: Path, extensions: Set[str] = None) -> tuple[int, int]:
    """
    Formatiert alle XML-Dateien in einem Verzeichnis rekursiv.
    
    Args:
        directory: Verzeichnis zum Durchsuchen
        extensions: Set von Dateiendungen (mit Punkt), Standard: {'.xml', '.rels'}
    
    Returns:
        tuple: (Anzahl erfolgreich formatiert, Anzahl gesamt)
    """
    if extensions is None:
        extensions = {'.xml', '.rels', '.vml'}
    
    total = 0
    success = 0
    
    # Alle Dateien rekursiv durchsuchen
    for file_path in directory.rglob('*'):
        if file_path.is_file() and file_path.suffix.lower() in extensions:
            total += 1
            if prettify_xml_file(file_path):
                success += 1
    
    return success, total


def extract_ooxml(file_path: Path, target_dir: Path, overwrite: bool, prettify: bool) -> Path:
    """
    Entpackt eine OOXML-Datei in den Zielordner.
    
    Args:
        file_path: Pfad zur OOXML-Datei
        target_dir: Zielordner für die Extraktion
        overwrite: Existierenden Ordner überschreiben
        prettify: XML-Dateien lesbar formatieren
    
    Returns:
        Path: Pfad zum erstellten Ordner
    """
    if not file_path.exists():
        raise click.ClickException(f"Datei nicht gefunden: {file_path}")
    
    if not zipfile.is_zipfile(file_path):
        raise click.ClickException(f"Datei ist kein gültiges ZIP/OOXML-Archiv: {file_path}")
    
    # Zielordner bestimmen
    if overwrite and target_dir.exists():
        click.echo(f"Lösche existierenden Ordner: {target_dir}")
        shutil.rmtree(target_dir)
        final_target = target_dir
    elif not overwrite and target_dir.exists():
        final_target = get_unique_folder_name(target_dir)
        click.echo(f"Ordner existiert bereits, verwende: {final_target.name}")
    else:
        final_target = target_dir
    
    # Ordner erstellen und entpacken
    final_target.mkdir(parents=True, exist_ok=False)
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(final_target)
        
        click.echo(f"✓ Erfolgreich entpackt nach: {final_target}")
        
        # Optional: XML-Dateien formatieren
        if prettify:
            click.echo("Formatiere XML-Dateien...")
            success, total = prettify_xml_files(final_target)
            click.echo(f"✓ {success} von {total} XML-Dateien formatiert")
        
        #if extract_vba_project(file_path, final_target / "vbaProject"):
        #    click.echo("✓ VBA-Projekt extrahiert nach: vbaProject/")
        if export_vba_project(file_path, final_target / "vbaProject"):
            click.echo("✓ VBA-Projekt extrahiert nach: vbaProject/")
        
        return final_target
    
    except Exception as e:
        # Bei Fehler aufräumen
        if final_target.exists():
            shutil.rmtree(final_target)
        raise click.ClickException(f"Fehler beim Entpacken: {e}")


def pack_ooxml(source_dir: Path, target_file: Path, overwrite: bool) -> Path:
    """
    Packt einen Ordner zu einer OOXML-Datei.
    XML-Dateien werden automatisch minimiert.
    
    Args:
        source_dir: Quellordner mit entpackten OOXML-Dateien
        target_file: Zieldatei (.xlsx, .docx, etc.)
        overwrite: Existierende Datei überschreiben
    
    Returns:
        Path: Pfad zur erstellten Datei
    """
    if not source_dir.exists():
        raise click.ClickException(f"Ordner nicht gefunden: {source_dir}")
    
    if not source_dir.is_dir():
        raise click.ClickException(f"Kein Ordner: {source_dir}")
    
    # Prüfen ob Zieldatei existiert
    if target_file.exists() and not overwrite:
        raise click.ClickException(
            f"Datei existiert bereits: {target_file}\n"
            f"Verwende --force zum Überschreiben."
        )
    
    # Parent-Verzeichnis erstellen falls nötig
    target_file.parent.mkdir(parents=True, exist_ok=True)
    
    # XML-Endungen die minimiert werden sollen
    xml_extensions = {'.xml', '.rels', '.vml'}
    
    xml_count = 0
    file_count = 0
    
    try:
        with zipfile.ZipFile(target_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Alle Dateien rekursiv durchgehen
            for file_path in sorted(source_dir.rglob('*')):
                if file_path.is_file():
                    # VBA-Projekt wurde nur zum lesen extrahiert und darf nicht ins OOXML-Archiv
                    if file_path.parent.name == 'vbaProject':
                        continue

                    #TODO: Update vbaProject.bin without Visio Application
                    #if file_path.name == 'vbaProject.bin':
                    #    click.echo("Aktualisiere vbaProject.bin... ACHTUNG: Das funktioniert nicht!")
                    #    update_vba_project_bin(file_path, file_path.parent.parent / 'vbaProject')

                    # Relativer Pfad im ZIP
                    arcname = file_path.relative_to(source_dir)
                    # XML-Dateien minimieren
                    if file_path.suffix.lower() in xml_extensions:
                        xml_data = minify_xml_file_to_bytes(file_path)
                        zipf.writestr(str(arcname), xml_data)
                        xml_count += 1
                    else:
                        # Andere Dateien direkt hinzufügen
                        zipf.write(file_path, arcname)
                    
                    file_count += 1
        
        success = import_vba_project(target_file, source_dir / 'vbaProject')

        click.echo(f"✓ {file_count} Dateien gepackt ({xml_count} XML-Dateien minimiert)")
        click.echo(f"✓ Datei erstellt: {target_file}")
        if success:
            click.echo("✓ VBA-Projekt aktualisiert in vbaProject.bin")
        else:
            click.echo("✗ Kein VBA-Projekt zum Aktualisieren gefunden.")  
        
        return target_file
    
    except Exception as e:
        # Bei Fehler aufräumen
        if target_file.exists():
            target_file.unlink()
        raise click.ClickException(f"Fehler beim Packen: {e}")
