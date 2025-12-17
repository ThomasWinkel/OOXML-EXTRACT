import click
from pathlib import Path
from .ooxml_package import extract_ooxml, pack_ooxml
from .ooxml_merge import automerge


@click.group()
@click.version_option(version="0.1.0")
def cli():
    """OOXML Extractor - Entpackt und packt Office-Dateien im OOXML-Format"""
    pass


@cli.command("extract")
@click.argument('file', type=click.Path(exists=True, path_type=Path))
@click.option(
    '-o', '--output',
    type=click.Path(path_type=Path),
    help='Zielordner für die Extraktion (Standard: Ordner mit Dateinamen)'
)
@click.option(
    '-f', '--force',
    is_flag=True,
    help='Existierenden Ordner ohne Rückfrage überschreiben'
)
@click.option(
    '-p', '--prettify',
    is_flag=True,
    help='XML-Dateien lesbar formatieren (mit Einrückung und Zeilenumbrüchen)'
)
def cli_extract(file: Path, output: Path, force: bool, prettify: bool):
    """
    Entpackt eine OOXML-Datei (xlsx, xlsm, vsdx, docx, pptx, etc.)
    
    Beispiele:
    
      ooxml extract dokument.xlsx
      
      ooxml extract dokument.xlsx -o /pfad/zum/zielordner
      
      ooxml extract dokument.xlsx --force --prettify
    """
    file = file.resolve()
    
    # Zielordner bestimmen
    if output:
        target_dir = output.resolve()
    else:
        # Standard: Ordner mit Dateinamen (ohne Endung) im selben Verzeichnis
        target_dir = file.parent / file.stem
    
    click.echo(f"Entpacke: {file.name}")
    
    extract_ooxml(file, target_dir, force, prettify)


@cli.command("pack")
@click.argument('directory', type=click.Path(exists=True, file_okay=False, path_type=Path))
@click.argument('output', type=click.Path(path_type=Path))
@click.option(
    '-f', '--force',
    is_flag=True,
    help='Existierende Datei ohne Rückfrage überschreiben'
)
def cli_pack(directory: Path, output: Path, force: bool):
    """
    Packt einen Ordner zu einer OOXML-Datei.
    XML-Dateien werden automatisch minimiert.
    
    Beispiele:
    
      ooxml pack dokument ausgabe.xlsx
      
      ooxml pack ./extracted/dokument repariert.xlsx
      
      ooxml pack dokument neu.xlsx --force
    """
    directory = directory.resolve()
    output = output.resolve()
    
    click.echo(f"Packe Ordner: {directory.name}")
    
    pack_ooxml(directory, output, force)


@cli.command("automerge")
@click.argument('ooxml_original', type=click.Path(exists=True, file_okay=True, dir_okay=False, path_type=Path))
@click.argument('ooxml_a', type=click.Path(exists=True, file_okay=True, dir_okay=False, path_type=Path))
@click.argument('ooxml_b', type=click.Path(exists=True, file_okay=True, dir_okay=False, path_type=Path))
@click.argument('ooxml_merged', type=click.Path(exists=False, file_okay=True, dir_okay=False, path_type=Path))
@click.option(
    '-f', '--force',
    is_flag=True,
    help='Overwrite existing file without prompt'
)
def cli_automerge(ooxml_original: Path, ooxml_a: Path, ooxml_b: Path, ooxml_merged: Path, force: bool):
    """
    Merges changes from two modified OOXML files based on an original file.
    
    Examples:
      ooxml automerge original.xlsx modified_a.xlsx modified_b.xlsx merged.xlsx
      ooxml automerge original.vsdx mod_a.vsdx mod_b.vsdx result.vsdx --force
    """
    automerge(ooxml_original, ooxml_a, ooxml_b, ooxml_merged, force)
