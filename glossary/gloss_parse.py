import json
from pathlib import Path
from csv import DictReader

import pyewts
from botok import TokChunks
from tibetan_sort import TibetanSort

converter = pyewts.pyewts()
sorter = TibetanSort()


def parse_glossaries(in_path, out_path):
    # parse txt glossaries
    joined = {}
    gloss_txt = Path(in_path) / 'raw_glossaries'
    for f in gloss_txt.glob('*.txt'):
        print(f.name)
        parse_bar_separated(f, joined)

    # todo: parse spreadsheet glossary
    gloss_csv = Path(in_path) / 'spreadsheets'
    for f in gloss_csv.glob('*.csv'):
        print(f.name)
        parse_csv(f, joined)

    # sort content
    total = {}
    sorted_words = sorter.sort_list(joined.keys())
    for num, word in enumerate(sorted_words):
        entry = joined[word]
        sorted_entry = [(k, entry[k]) for k in sorted(entry.keys())]
        total[num+1] = (word, sorted_entry)

    out_file = Path(out_path) / 'glossary.json'
    out_file.write_text(json.dumps(total, ensure_ascii=False), encoding='utf8')


def parse_csv(in_file, joined):
    with in_file.open(newline='') as csvfile:
        reader = DictReader(csvfile)
        for row in reader:
            word = row['Tibétain'].strip()
            if not word or word[0].isupper():
                continue
            others = ''
            if '\n' in word:
                word, others = word.split('\n', maxsplit=1)
            others = others.strip()
            word = converter.toUnicode(word)
            word = '་'.join(TokChunks(word).get_syls())
            word = add_shad(word)
            skrt = row['Sanskrit'].strip()
            root_en = row['Anglais'].strip()
            root_fr = row['Sens racine Français'].strip()
            alter = row['Autres termes rencontrés'].strip()
            entry = []
            if others:
                entry.append(others)
            if skrt:
                entry.append(skrt)
            roots = [r for r in [root_fr, root_en] if r]
            if roots:
                entry.append(f'Sens racine: {", ".join(roots)}')
            if alter:
                entry.append(f'Alternatives: {alter}')
            entry = '\n'.join(entry)

            if word not in joined:
                joined[word] = {}
            joined[word][in_file.stem] = [entry]


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
