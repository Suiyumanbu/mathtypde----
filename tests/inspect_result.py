from __future__ import annotations

from pathlib import Path
import argparse
import json
import re
import zipfile


parser = argparse.ArgumentParser()
parser.add_argument('--output', type=Path, default=Path(__file__).resolve().parent / 'output')
root = parser.parse_args().output.resolve()


def inspect_docx(name: str) -> dict:
    path = root / name
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        return {
            'xml': archive.read('word/document.xml').decode('utf-8'),
            'custom': (
                archive.read('docProps/custom.xml').decode('utf-8')
                if 'docProps/custom.xml' in names
                else ''
            ),
            'embeddings': {
                entry: archive.read(entry)
                for entry in names
                if entry.startswith('word/embeddings/')
            },
        }


def read_report(name: str) -> dict:
    return json.loads((root / name).read_text(encoding='utf-8-sig'))


def equation_sequence_results(document_xml: str) -> list[int]:
    values: list[int] = []
    for paragraph in re.findall(r'<w:p\b.*?</w:p>', document_xml):
        if 'MACROBUTTON MTPlaceRef' not in paragraph:
            continue
        marker = paragraph.rfind(r'SEQ MTEqn \c')
        assert marker >= 0, paragraph
        match = re.search(r'<w:instrText>(\d+)</w:instrText>', paragraph[marker:])
        assert match, paragraph
        values.append(int(match.group(1)))
    return values


initial = inspect_docx('result.docx')
initial_report = read_report('result.report.json')
assert len(initial_report['equations']) == 7
assert all(item['progId'] == 'Equation.DSMT4' for item in initial_report['equations'])
assert all(
    item['nativeNumberFields'] > 0
    for item in initial_report['equations']
    if item['mode'] == 'right-numbered'
)
assert len(initial['embeddings']) == 7, initial['embeddings'].keys()
assert initial['xml'].count('<o:OLEObject') == 7
assert initial['xml'].count('MACROBUTTON MTPlaceRef') == 2
assert initial['xml'].count('MTDisplayEquation') == 4
assert '[[EQ:' not in initial['xml']
assert 'UNCHANGED_FIELD' in initial['xml'], 'An unrelated AUTHOR field was updated.'
assert 'MTDeferFieldUpdate' not in initial['custom'], 'Temporary field-update setting leaked.'
assert 'MTEquationSection' in initial['custom'], 'MathType numbering section metadata is missing.'
assert 'mt_bookmark_display' in initial['xml'], 'The reusable display-equation bookmark was lost.'
assert initial_report['timings']['processingMilliseconds'] > 0

renumbered = inspect_docx('renumbered.docx')
renumbered_report = read_report('renumbered.report.json')
assert len(renumbered_report['equations']) == 2
assert all(item['operation'] == 'number-existing' for item in renumbered_report['equations'])
assert all(item['source'] == 'existing' for item in renumbered_report['equations'])
assert all(item['mode'] == 'right-numbered' for item in renumbered_report['equations'])
assert all(item['progId'] == 'Equation.DSMT4' for item in renumbered_report['equations'])
assert all(item['nativeNumberFields'] > 0 for item in renumbered_report['equations'])
assert all(item['conversionMilliseconds'] == 0 for item in renumbered_report['equations'])
assert len(renumbered['embeddings']) == 7
assert renumbered['embeddings'] == initial['embeddings'], 'Existing OLE payloads were recreated.'
assert renumbered['xml'].count('<o:OLEObject') == 7
assert renumbered['xml'].count('MACROBUTTON MTPlaceRef') == 4
assert renumbered['xml'].count('MTDisplayEquation') == 4
assert equation_sequence_results(renumbered['xml']) == [1, 2, 3, 4]
assert 'UNCHANGED_FIELD' in renumbered['xml'], 'An unrelated AUTHOR field was updated.'
assert 'MTDeferFieldUpdate' not in renumbered['custom'], 'Temporary field-update setting leaked.'
assert 'mt_bookmark_display' in renumbered['xml'], 'The existing-equation bookmark was lost.'
assert renumbered_report['nativeNumberParagraphsUpdated'] == 4

print(json.dumps({
    'initialEquations': len(initial_report['equations']),
    'initialNativeNumbered': initial['xml'].count('MACROBUTTON MTPlaceRef'),
    'existingEquationsNumbered': len(renumbered_report['equations']),
    'finalNativeNumbered': renumbered['xml'].count('MACROBUTTON MTPlaceRef'),
    'olePayloadsReused': True,
    'sequenceResults': equation_sequence_results(renumbered['xml']),
    'initialProcessingMilliseconds': initial_report['timings']['processingMilliseconds'],
    'numberExistingProcessingMilliseconds': renumbered_report['timings']['processingMilliseconds'],
}, indent=2))
