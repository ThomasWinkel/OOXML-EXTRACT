#from oletools.olevba import VBA_Parser
import win32com.client as win32
from pathlib import Path


def export_vba_project(vsdm_file: Path, vba_files_dir: Path) -> bool:
    vsdm_file = Path(vsdm_file).resolve()
    vba_files_dir = Path(vba_files_dir).resolve()
    
    if not vsdm_file.exists():
        return False

    visio_app = None
    document = None
    success = False

    try:
        visio_app = win32.Dispatch("Visio.Application")
        visio_app.Visible = False
        document = visio_app.Documents.OpenEx(str(vsdm_file), 128 + 256) # visOpenMacrosDisabled + visOpenNoWorkspace
        vb_project = document.VBProject
        
        vba_files_dir.mkdir(parents=True, exist_ok=True)
        
        for component in vb_project.VBComponents:
            if component.Type == 100:
                file_name = component.Name + '.cls'
                export_path = vba_files_dir / file_name
                code_module = component.CodeModule
                count_of_lines = code_module.CountOfLines
                lines = ''
                if count_of_lines > 0:
                    lines = code_module.Lines(1, count_of_lines)
                with open(export_path, 'w', encoding='utf-8', newline='') as f:
                    f.write(lines)
                print(f"Exportiert: {export_path}")
            elif component.Type == 1:
                file_name = component.Name + '.bas'
                export_path = vba_files_dir / file_name
                component.Export(str(export_path))
                print(f"Exportiert: {export_path}")
            elif component.Type == 2:
                file_name = component.Name + '.cls'
                export_path = vba_files_dir / file_name
                component.Export(str(export_path))
                print(f"Exportiert: {export_path}")
            elif component.Type == 3:
                file_name = component.Name + '.frm'
                export_path = vba_files_dir / file_name
                component.Export(str(export_path))
                print(f"Exportiert: {export_path}")
        
        success = True
        
    except Exception as e:
        print(f"Ein schwerwiegender Fehler ist aufgetreten: {e}")
        success = False
        
    finally:
        if document:
            try:
                document.Saved = True
                document.Close() 
            except:
                pass 

        if visio_app:
            visio_app.Quit()
            
    return success


def import_vba_project(vsdm_file: Path, vba_files_dir: Path) -> bool:
    vsdm_file = Path(vsdm_file).resolve()
    vba_files_dir = Path(vba_files_dir).resolve()
    
    if not vsdm_file.exists() or not vba_files_dir.is_dir():
        return False

    visio_app = None
    document = None
    success = False

    try:
        visio_app = win32.Dispatch("Visio.Application")
        visio_app.Visible = False
        document = visio_app.Documents.OpenEx(str(vsdm_file), 128 + 256) # visOpenMacrosDisabled + visOpenNoWorkspace
        vb_project = document.VBProject
        
        for component in vb_project.VBComponents:
            if component.Type in [1, 2, 3]:
                vb_project.VBComponents.Remove(component)
            elif component.Type == 100:
                code_module = component.CodeModule
                count_of_lines = code_module.CountOfLines
                if count_of_lines > 0:
                    code_module.DeleteLines(1, count_of_lines)
        
        vba_files = [f for f in vba_files_dir.iterdir() if f.suffix.lower() in ['.bas', '.cls', '.frm']]
        
        print(f"Importiere {len(vba_files)} neue Komponenten...")
        for file in vba_files:
            try:
                if file.name.lower() == 'ThisDocument.cls'.lower():
                    code_module.AddFromFile(str(file))
                else:
                    vb_project.VBComponents.Import(str(file))
                print(f"  - Importiert: {file.name}")
            except Exception as e:
                print(f"  - Fehler beim Importieren von {file.name}: {e}")

        document.Save()
        success = True
        
    except Exception as e:
        print(f"Ein schwerwiegender Fehler ist aufgetreten: {e}")
        success = False
        
    finally:
        if document:
            try:
                document.Saved = True
                document.Close() 
            except:
                pass 

        if visio_app:
            visio_app.Quit()
            
    return success


# Implementierung mit oletools, ohne COM. Export geht, Import/Update von vbaProject.bin ist zu kompliziert.
# def extract_vba_project(ooxml_path: Path, output_dir: Path) -> bool:
#     vbaparser = VBA_Parser(ooxml_path)
    
#     if not vbaparser.detect_vba_macros():
#         return False
    
#     output_dir.mkdir(parents=True, exist_ok=True)
    
#     for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
#         module_file = output_dir / f"{vba_filename}"
#         with open(module_file, 'w', encoding='utf-8', newline='',  errors='ignore') as f:
#             f.write(vba_code)
            
#     vbaparser.close()
#     return True


# def update_vba_project_bin(vba_project_bin: Path, vba_files: Path) -> bool:
#     """
#     Aktualisiert den Code in vbaProject.bin.
#     TODO: Das scheint sehr schwierig zu sein. Das lasse ich erstmal weg.
#     """
#     pass
