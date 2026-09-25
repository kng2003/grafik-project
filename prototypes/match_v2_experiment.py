# -*- coding: utf-8 -*-
"""
Эксперимент: сравнить текущий метод (Jaccard по словоформам) с версией,
где слова приведены к начальной форме (pymorphy2) и добавлено
TF-IDF + косинусное сходство (sklearn), чтобы честно ответить на вопрос
"поднимет ли установка библиотек точность".
"""
import re, glob, sys, time
sys.path.insert(0, 'prototypes')
from parse_frsn_prototype import parse_file
import openpyxl
import pymorphy2
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

t0 = time.time()
morph = pymorphy2.MorphAnalyzer()
_lemma_cache = {}
def lemma(word):
    v = _lemma_cache.get(word)
    if v is None:
        v = morph.parse(word)[0].normal_form
        _lemma_cache[word] = v
    return v

STOPWORDS_LEMMA = set("""
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
    return [w for w in lemmas if w not in STOPWORDS_LEMMA]

def numbers(s):
    return set(re.findall(r'\d+[.,]?\d*', s or ''))

# старая (без лемм) функция для сравнения
def norm_tokens_old(s):
    STOP_OLD = set("""
    и в с по от до на из для не или так как при над под без через между также ещё эту это тот та то те
    диаметр диаметром толщина толщиной устройство работы работ мм м2 м3 шт кг см номинальным
    группа группой при выполнении высотой высота
    """.split())
    s = (s or '').lower()
    s = re.sub(r'[«»"\'()]', ' ', s)
    s = re.sub(r'[.,:;]', ' ', s)
    words = re.findall(r'[а-яё]{3,}', s)
    return [w for w in words if w not in STOP_OLD]

print(f"[{time.time()-t0:.1f}s] старт")

wb = openpyxl.load_workbook("prototypes/results/2026-09-16_gesn_codes_full_list.xlsx", read_only=True, data_only=True)
ws = wb["Коды ГЭСН"]
gesn = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row and row[1]:
        gesn.append({'code': row[1], 'name': row[2], 'unit': row[3]})
print(f"[{time.time()-t0:.1f}s] кодов ГЭСН: {len(gesn)}")

files = sorted(glob.glob("docs/ФРСН/*.md"))
frsn_records = []
for path in files:
    m = re.match(r'docs/ФРСН/(\d+) ', path)
    sbornik_num = int(m.group(1)) if m else -1
    recs = parse_file(path)
    for r in recs:
        if not r['valid_shifr']:
            continue
        r['sbornik'] = sbornik_num
        frsn_records.append(r)
print(f"[{time.time()-t0:.1f}s] записей ФРСН: {len(frsn_records)}")

# === Метод СТАРЫЙ (Jaccard, без лемм) ===
index_old = {}
rec_tokens_old = []
for i, r in enumerate(frsn_records):
    toks = set(norm_tokens_old(r['name']))
    rec_tokens_old.append(toks)
    for w in toks:
        index_old.setdefault(w, []).append(i)

def score_old(g):
    q = set(norm_tokens_old(g['name']))
    qn = numbers(g['name'])
    is_m = g['code'].upper().startswith('ГЭСНМ')
    cand = set()
    for w in q:
        cand.update(index_old.get(w, []))
    best = 0.0
    for i in cand:
        r = frsn_records[i]
        rt = rec_tokens_old[i]
        if not q or not rt:
            continue
        jacc = len(q & rt) / len(q | rt)
        rn = numbers(r['name'])
        nb = 0.25 if (qn and rn and qn & rn) else 0.0
        fb = 0.15 if (is_m and 49 <= r['sbornik'] <= 72) else (-0.1 if is_m else 0.0)
        sc = jacc + nb + fb
        if sc > best:
            best = sc
    return best

print(f"[{time.time()-t0:.1f}s] считаю старый метод...")
old_scores = [score_old(g) for g in gesn]
print(f"[{time.time()-t0:.1f}s] старый метод готов")

# === Метод НОВЫЙ (леммы + Jaccard) ===
index_new = {}
rec_lemmas = []
for i, r in enumerate(frsn_records):
    lm = set(norm_lemmas(r['name']))
    rec_lemmas.append(lm)
    for w in lm:
        index_new.setdefault(w, []).append(i)
print(f"[{time.time()-t0:.1f}s] леммы ФРСН построены, слов: {len(index_new)}")

def score_new_jaccard(g):
    q = set(norm_lemmas(g['name']))
    qn = numbers(g['name'])
    is_m = g['code'].upper().startswith('ГЭСНМ')
    cand = set()
    for w in q:
        cand.update(index_new.get(w, []))
    best = 0.0
    for i in cand:
        r = frsn_records[i]
        rt = rec_lemmas[i]
        if not q or not rt:
            continue
        jacc = len(q & rt) / len(q | rt)
        rn = numbers(r['name'])
        nb = 0.25 if (qn and rn and qn & rn) else 0.0
        fb = 0.15 if (is_m and 49 <= r['sbornik'] <= 72) else (-0.1 if is_m else 0.0)
        sc = jacc + nb + fb
        if sc > best:
            best = sc
    return best

print(f"[{time.time()-t0:.1f}s] считаю метод с леммами...")
new_scores = [score_new_jaccard(g) for g in gesn]
print(f"[{time.time()-t0:.1f}s] метод с леммами готов")

# === Метод TF-IDF + косинус (на леммах, sklearn) ===
corpus_frsn = [' '.join(sorted(rec_lemmas[i])) for i in range(len(frsn_records))]
corpus_gesn = [' '.join(sorted(norm_lemmas(g['name']))) for g in gesn]
vec = TfidfVectorizer()
X_frsn = vec.fit_transform(corpus_frsn)
X_gesn = vec.transform(corpus_gesn)
print(f"[{time.time()-t0:.1f}s] TF-IDF векторизация готова, признаков: {len(vec.vocabulary_)}")

tfidf_scores = []
BATCH = 20
for start in range(0, len(gesn), BATCH):
    end = min(start+BATCH, len(gesn))
    sims = cosine_similarity(X_gesn[start:end], X_frsn)
    for row in sims:
        tfidf_scores.append(float(row.max()) if row.size else 0.0)
print(f"[{time.time()-t0:.1f}s] TF-IDF скоры готовы")

def stats(name, arr):
    weak = sum(1 for x in arr if x < 0.3)
    print(f"{name}: среднее={sum(arr)/len(arr):.3f}  мин={min(arr):.3f}  макс={max(arr):.3f}  слабых(<0.3)={weak}/{len(arr)}")

stats("СТАРЫЙ (словоформы, Jaccard)      ", old_scores)
stats("НОВЫЙ  (леммы pymorphy2, Jaccard) ", new_scores)
stats("TF-IDF (леммы, cosine, sklearn)   ", tfidf_scores)

# точечная проверка контрольного примера
for g, so, sn in zip(gesn, old_scores, new_scores):
    if 'ГЭСН22-03-007-02' in str(g['code']):
        print(f"\nКонтрольный пример {g['code']}: старый скор={so:.3f}, новый скор(леммы)={sn:.3f}")

improved = sum(1 for o, n in zip(old_scores, new_scores) if n > o + 0.01)
worse = sum(1 for o, n in zip(old_scores, new_scores) if n < o - 0.01)
same = len(gesn) - improved - worse
print(f"\nЛеммы против старого метода: улучшилось {improved}, ухудшилось {worse}, без изменений {same} (из {len(gesn)})")
