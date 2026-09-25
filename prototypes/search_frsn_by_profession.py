# -*- coding: utf-8 -*-
import re, glob, sys
sys.path.insert(0, '/home/user/grafik-project/prototypes')
from parse_frsn_prototype import parse_file
import pymorphy3
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

morph = pymorphy3.MorphAnalyzer()
_cache = {}
def lemma(w):
    v = _cache.get(w)
    if v is None:
        v = morph.parse(w)[0].normal_form
        _cache[w] = v
    return v

STOPWORDS = set("""
и в с по от до на из для не или так как при над под без через между также ещё этот тот то те
диаметр толщина устройство работа мм м2 м3 шт кг см номинальный
группа при выполнение высота
""".split())

def norm_lemmas(s):
    s = (s or '').lower()
    s = re.sub(r'[«»"\'()]', ' ', s)
    s = re.sub(r'[.,:;]', ' ', s)
    words = re.findall(r'[а-яё]{3,}', s)
    lemmas = [lemma(w) for w in words]
    return [w for w in lemmas if w not in STOPWORDS]

FILES = sorted(glob.glob('/home/user/grafik-project/docs/ФРСН/ФРСН_md/*.md'))
records = []
for path in FILES:
    m = re.match(r'.*/ФРСН_md/(\d+) ', path)
    sbornik = int(m.group(1)) if m else -1
    for r in parse_file(path):
        if not r['valid_shifr']:
            continue
        r['sbornik'] = sbornik
        records.append(r)

lemma_strs = [' '.join(sorted(norm_lemmas(r['name']))) for r in records]
vec = TfidfVectorizer()
X_all = vec.fit_transform(lemma_strs)
print(f"Records: {len(records)}", file=sys.stderr)

def crew_str(r):
    if not r['crew']:
        return '(нет)'
    return "; ".join(f"{c['profession']} x{c['qty']}" for c in r['crew'])

def has_profession(r, kws):
    profs = ' '.join(c['profession'].lower() for c in r['crew'])
    return any(kw.lower() in profs for kw in kws)

def search(query_name, profession_kw=None, sbornik_filter=None, top=8):
    idx = list(range(len(records)))
    if sbornik_filter:
        idx = [i for i in idx if records[i]['sbornik'] in sbornik_filter]
    if profession_kw:
        idx = [i for i in idx if has_profession(records[i], profession_kw)]
    if not idx:
        return []
    qv = vec.transform([' '.join(sorted(norm_lemmas(query_name)))])
    sims = cosine_similarity(qv, X_all[idx])[0]
    scored = sorted(zip(sims, idx), key=lambda x: -x[0])[:top]
    return [(s, records[i]) for s, i in scored]

def show(results):
    for score, r in results:
        print(f"  [{score:.3f}] {r['shifr']}  сб.{r['sbornik']:02d}  | {r['name']}  | чел-ч {r['lab_hours']} | {crew_str(r)}")
