import xml.etree.ElementTree as ET
import zipfile

from openpyxl import Workbook

from sow_merge_tool.sheet_screening import NS, sheet_fingerprints


def pair(tmp_path, transform):
    left, right = tmp_path / 'before.xlsx', tmp_path / 'after.xlsx'
    wb = Workbook()
    wb.active['A1'] = 'original'
    wb.active['B2'] = '=1+1'
    wb.save(left)
    wb.close()
    with zipfile.ZipFile(left) as source, zipfile.ZipFile(right, 'w') as dest:
        for name in source.namelist():
            dest.writestr(name, transform(name, source.read(name)))
    return left, right


def test_view_changes_are_skipped_but_values_and_formulas_are_not(tmp_path):
    def viewport(name, data):
        if name == 'xl/worksheets/sheet1.xml':
            root = ET.fromstring(data)
            root.find(NS + 'sheetViews/' + NS + 'sheetView').set('topLeftCell', 'Z200')
            return ET.tostring(root)
        return data
    left, right = pair(tmp_path, viewport)
    assert sheet_fingerprints(left) == sheet_fingerprints(right)
    for before, after in [(b'original', b'changed'), (b'1+1', b'1+2')]:
        left, right = pair(tmp_path, lambda name, data, before=before, after=after: data.replace(before, after))
        assert sheet_fingerprints(left) != sheet_fingerprints(right)


def test_style_dependency_change_is_not_hidden(tmp_path):
    left, right = pair(tmp_path, lambda name, data: data.replace(b'Calibri', b'Arial')
                       if name == 'xl/styles.xml' else data)
    assert sheet_fingerprints(left) != sheet_fingerprints(right)


def test_shared_string_change_with_identical_sheet_xml(tmp_path):
    left, right = tmp_path / 'a.xlsx', tmp_path / 'b.xlsx'
    for path, value in [(left, 'old'), (right, 'new')]:
        with zipfile.ZipFile(path, 'w') as z:
            z.writestr('xl/workbook.xml', f'<workbook xmlns="{NS[1:-1]}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" r:id="r1"/></sheets></workbook>')
            z.writestr('xl/_rels/workbook.xml.rels', '<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>')
            z.writestr('xl/worksheets/sheet1.xml', f'<worksheet xmlns="{NS[1:-1]}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>')
            z.writestr('xl/sharedStrings.xml', f'<sst xmlns="{NS[1:-1]}"><si><t>{value}</t></si></sst>')
    assert sheet_fingerprints(left) != sheet_fingerprints(right)
