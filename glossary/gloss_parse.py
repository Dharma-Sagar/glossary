from pathlib import Path

from docx import Document
import pyewts
from botok import TokChunks
from tibetan_sort import TibetanSort

converter = pyewts.pyewts()
sorter = TibetanSort()


def parse_glossaries(in_path, out_file):
    # parse txt glossaries
    joined = {}
    gloss_txt = Path(in_path) / 'raw_glossaries'
    for f in gloss_txt.glob('*.txt'):
        print(f.name)
        parse_bar_separated(f, joined)

    # todo: parse spreadsheet glossary

    # sort content
    total = []
    sorted_words = sorter.sort_list(joined.keys())
    for num, word in enumerate(sorted_words):
        entry = joined[word]
        sorted_entry = [(k, entry[k]) for k in sorted(entry.keys())]
        total.append((f'{num+1} {word}', sorted_entry))

    # export in docx
    export_docx(total, out_file)


def parse_bar_separated(in_file, joined):
    sep = '|'
    dump = in_file.read_text()
    name = in_file.stem

    # sanity check
    mal_formed = []

    for num, line in enumerate(dump.split('\n')):
        if not line.strip():
            continue

        if line.count(sep) != 1:
            mal_formed.append((num+1, line))
        else:
            word, defnt = [e.strip() for e in line.split(sep)]

            # cleanup (TokChunks) and convert to unicode
            word = converter.toUnicode(word)
            word = '་'.join(TokChunks(word).get_syls())
            word = add_shad(word)

            if word not in joined:
                joined[word] = {}

            if name not in joined[word]:
                joined[word][name] = []

            joined[word][name].append(defnt)
    if mal_formed:
        print('\n'.join([': '.join(m) for m in mal_formed]))
        exit(1)


def add_shad(word):
    if word.endswith('ང'):
        return word + '་།'
    elif word[-1] in ['ཀ', 'ག', 'ཤ']:
        return word
    else:
        return word + '།'


def export_docx(data, out_file):
    doc = Document()

    for word, entries in data:
        doc.add_heading(word, 3)
        for name, entry in entries:
            doc.add_heading(name, 5)
            p = doc.add_paragraph()
            for n, defnt in enumerate(entry):
                if n > 0:
                    p.add_run().add_break()
                p.add_run(defnt)

    doc.save(out_file)


parse_glossaries('../content', '../content/Glossaire.docx')
