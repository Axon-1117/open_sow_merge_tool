"""Conservative worksheet equality proof before expensive alignment."""
from __future__ import annotations

import hashlib
import posixpath
import xml.etree.ElementTree as ET
import zipfile

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def sheet_fingerprints(path: str, *, formula_flags: dict | None = None) -> dict[str, str]:
    """Ignore saved viewport/dimension hints, retain content and dependencies.

    Missing or unsupported evidence is never equality. Shared string indices
    are resolved so a different string table cannot hide a cell change.
    """
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        if len(names) != len(set(names)):
            raise ValueError('Duplicate workbook package members')
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        relations = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'): r.get('Target', '') for r in relations}
        strings = []
        if 'xl/sharedStrings.xml' in names:
            strings = [hashlib.sha256(ET.tostring(item)).hexdigest()
                       for item in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        dependencies = hashlib.sha256()
        for name in sorted(names):
            if (name.startswith('xl/') and not name.startswith('xl/worksheets/')
                    and name not in {'xl/workbook.xml', 'xl/sharedStrings.xml',
                                     'xl/calcChain.xml'}):
                dependencies.update(name.encode())
                dependencies.update(archive.read(name))
        for parent in workbook.iter():
            for child in list(parent):
                if child.tag == '{http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac}absPath':
                    parent.remove(child)
                elif child.tag == NS + 'ext' and child.get('uri') == '{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}':
                    parent.remove(child)  # Excel producer calculation capabilities
        for child in workbook:
            if child.tag == NS + 'extLst' and not len(child):
                continue
            if child.tag not in {NS + 'sheets', NS + 'bookViews', NS + 'fileVersion',
                                 '{http://schemas.microsoft.com/office/spreadsheetml/2014/revision}revisionPtr'}:
                dependencies.update(ET.tostring(child))
        result = {}
        for sheet in workbook.findall(NS + 'sheets/' + NS + 'sheet'):
            rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target = targets.get(rid, '')
            part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
            if not part.startswith('xl/worksheets/'):
                continue
            root = ET.fromstring(archive.read(part))
            if formula_flags is not None:
                formula_flags[sheet.attrib['name']] = next(root.iter(NS + 'f'), None) is not None
            for tag in ('sheetViews', 'dimension'):
                for child in root.findall(NS + tag):
                    root.remove(child)
            for cell in root.iter(NS + 'c'):
                if cell.get('t') == 's':
                    value = cell.find(NS + 'v')
                    if value is None or value.text is None:
                        raise ValueError('Missing shared string reference')
                    index = int(value.text)
                    if index < 0 or index >= len(strings):
                        raise ValueError('Invalid shared string reference')
                    value.text = strings[index]
            digest = dependencies.copy()
            digest.update(sheet.get('state', 'visible').encode())
            digest.update(ET.tostring(root))
            # Relationships may reference comments/tables/drawings. Keep all
            # worksheet relationship changes conservative as well.
            relpart = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
            if relpart in names:
                digest.update(archive.read(relpart))
            result[sheet.attrib['name']] = digest.hexdigest()
        return result
