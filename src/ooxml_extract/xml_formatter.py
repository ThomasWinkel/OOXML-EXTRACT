import re
from pathlib import Path


def prettify_xml(xml_str: str, indent: str = "  ") -> str:
    match = re.match(r'<\?xml [^?]+\?>\s*', xml_str)
    
    if match:
        declaration = match.group(0).strip()
        xml_body = xml_str[len(match.group(0)):].strip()
    else:
        declaration = ""
        xml_body = xml_str.strip()

    pretty_str = re.sub(r'(?<=>)\s*(?=<)', r'\n', xml_body.strip())
    pretty_str = pretty_str.lstrip()
    formatted_lines = []
    level = 0
    
    for line in pretty_str.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith('</'):
            level -= 1
            formatted_lines.append(indent * level + line)
        elif line.startswith('<'):
            formatted_lines.append(indent * level + line)
            if not line.endswith('/>') and '</' not in line:
                level += 1
        else:
            formatted_lines.append(line)
    
    if declaration:
        return declaration + '\n' + '\n'.join(formatted_lines)
    else:
        return '\n'.join(formatted_lines)


def prettify_xml_file(file_path: Path) -> bool:
    """
    Formatiert eine XML-Datei und überschreibt sie mit der formatierten Version.
    
    Args:
        file_path: Pfad zur XML-Datei
        
    Returns:
        True bei Erfolg, False bei Fehler
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            xml_str = f.read()
        
        pretty_xml = prettify_xml(xml_str)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(pretty_xml)
        
        return True
        
    except Exception as e:
        # Optional: Logging des Fehlers
        # print(f"Fehler beim Formatieren von {file_path}: {e}")
        return False


def minify_xml(pretty_xml_str: str) -> str:
    match = re.match(r'<\?xml [^?]+\?>\s*', pretty_xml_str)

    if match:
        declaration = match.group(0).strip()
        xml_body = pretty_xml_str[len(match.group(0)):].strip()
    else:
        declaration = ""
        xml_body = pretty_xml_str.strip()

    minified_body = re.sub(r'>\s+<', '><', xml_body)

    return declaration + '\n' + minified_body


def minify_xml_file_to_bytes(file_path: Path) -> bytes:
    """
    Liest eine XML-Datei, minifiziert sie und gibt das Ergebnis als Bytes zurück.
    
    Args:
        file_path: Pfad zur XML-Datei
        
    Returns:
        Minifiziertes XML als Bytes
    """
    # XML-Datei einlesen
    with open(file_path, 'r', encoding='utf-8') as f:
        pretty_xml_str = f.read()
    
    # XML minifizieren
    minified_xml = minify_xml(pretty_xml_str)
    
    # In Bytes konvertieren und zurückgeben
    return minified_xml.encode('utf-8')
