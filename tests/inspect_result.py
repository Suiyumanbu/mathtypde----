from pathlib import Path
import json
import zipfile

root = Path(__file__).resolve().parent / 'output'
docx_path = root / 'result.docx'
report_path = root / 'result.report.json'

with zipfile.ZipFile(docx_path) as archive:
    names = archive.namelist()
    document_xml = archive.read('word/document.xml').decode('utf-8')
    custom_xml = (
        archive.read('docProps/custom.xml').decode('utf-8')
        if 'docProps/custom.xml' in names
        else ''
    )

report = json.loads(report_path.read_text(encoding='utf-8-sig'))
embeddings = [name for name in names if name.startswith('word/embeddings/')]

assert len(report['equations']) == 6
assert all(item['progId'] == 'Equation.DSMT4' for item in report['equations'])
assert all(item['nativeNumberFields'] > 0 for item in report['equations'] if item['mode'] == 'right-numbered')
assert len(embeddings) == 6, embeddings
assert document_xml.count('<o:OLEObject') == 6
assert document_xml.count('MACROBUTTON MTPlaceRef') == 2
assert document_xml.count('MTDisplayEquation') == 3
assert '[[EQ:' not in document_xml
assert 'UNCHANGED_FIELD' in document_xml, 'An unrelated AUTHOR field was updated.'
assert 'MTDeferFieldUpdate' not in custom_xml, 'Temporary field-update setting leaked into output.'
assert 'MTEquationSection' in custom_xml, 'MathType numbering section metadata is missing.'
print(json.dumps({
    'equations': len(report['equations']),
    'embeddings': len(embeddings),
    'nativeNumbered': document_xml.count('MACROBUTTON MTPlaceRef'),
    'displayStyleUses': document_xml.count('MTDisplayEquation'),
    'unrelatedFieldPreserved': True,
    'temporarySettingRestored': True,
}, indent=2))
