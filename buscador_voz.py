"""
Motor de búsqueda FONÉTICO para productos dictados por voz.

Qué resuelve (la base usa abreviaturas y marcas que ningún modelo de voz conoce):
  - "cuaderno"  -> CUAD      (abreviatura = comienzo de la palabra)
  - "estandar"  -> STD       (abreviatura = letras de la palabra en orden)
  - "hojas"     -> H         (alias del archivo alias.txt)
  - "kinota", "cuenota", "un nota" -> QNOTA   (suena parecido; se unen palabras vecinas)
No usa librerías externas.
"""
import math
import os
import re
import unicodedata
from collections import defaultdict

VOCALES = set("AEIOU")

# hablado -> abreviatura/palabra tal como está en la base (se amplía con alias.txt)
ALIAS_BASE = {
    "hojas": ["H"], "hoja": ["H"], "por": ["X"],
    "centimetros": ["CM"], "centimetro": ["CM"],
    "milimetros": ["MM"], "milimetro": ["MM"],
    "cuaderno": ["CUAD"], "cuadernos": ["CUAD"],
    "unidades": ["UND", "UN"], "unidad": ["UND", "UN"],
}
RELLENO = {"EL", "LA", "LOS", "LAS", "UN", "UNA", "UNOS", "UNAS", "DE", "Y", "CON", "PARA"}


def _sin_tildes(t):
    t = unicodedata.normalize("NFD", str(t).upper())
    return "".join(c for c in t if unicodedata.category(c) != "Mn")


_NUM = {"cero": 0, "un": 1, "uno": 1, "una": 1, "dos": 2, "tres": 3, "cuatro": 4, "cinco": 5,
        "seis": 6, "siete": 7, "ocho": 8, "nueve": 9, "diez": 10, "once": 11, "doce": 12,
        "trece": 13, "catorce": 14, "quince": 15, "dieciseis": 16, "diecisiete": 17,
        "dieciocho": 18, "diecinueve": 19, "veinte": 20, "veintiun": 21, "veintiuno": 21,
        "veintiuna": 21, "veintidos": 22, "veintitres": 23, "veinticuatro": 24,
        "veinticinco": 25, "veintiseis": 26, "veintisiete": 27, "veintiocho": 28,
        "veintinueve": 29, "treinta": 30, "cuarenta": 40, "cincuenta": 50, "sesenta": 60,
        "setenta": 70, "ochenta": 80, "noventa": 90, "cien": 100, "ciento": 100,
        "doscientos": 200, "doscientas": 200, "trescientos": 300, "trescientas": 300,
        "cuatrocientos": 400, "cuatrocientas": 400, "quinientos": 500, "quinientas": 500,
        "seiscientos": 600, "seiscientas": 600, "setecientos": 700, "setecientas": 700,
        "ochocientos": 800, "ochocientas": 800, "novecientos": 900, "novecientas": 900}
_DECENAS = {"treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"}


def numeros_a_cifras(texto):
    """'tres mil ochocientos' -> '3800' (los modelos de voz a veces escriben números con letras)."""
    t = _sin_tildes(texto).lower()
    pal, out, i = t.split(), [], 0
    es_num = lambda w: w in _NUM or w == "mil"
    while i < len(pal):
        if not es_num(pal[i]):
            out.append(pal[i])
            i += 1
            continue
        j, run = i, []
        while j < len(pal):
            w = pal[j]
            if es_num(w):
                run.append(w)
                j += 1
            elif (w == "y" and run and run[-1] in _DECENAS and j + 1 < len(pal)
                  and pal[j + 1] in _NUM):
                j += 1
            else:
                break
        if run in (["un"], ["una"]):
            out.append(pal[i])
        else:
            total = cur = 0
            for w in run:
                if w == "mil":
                    total += (cur or 1) * 1000
                    cur = 0
                else:
                    cur += _NUM[w]
            out.append(str(total + cur))
        i = j
    return " ".join(out)


def tokens(texto):
    t = _sin_tildes(numeros_a_cifras(texto))
    t = re.sub(r"(?<=\d)[.,](?=\d{3}(?!\d))", "", t)           # 19.000 = 19000
    t = re.sub(r"(?<=[A-Z])(?=\d)|(?<=\d)(?=[A-Z])", " ", t)   # 100H = 100 H
    t = re.sub(r"[^A-Z0-9]+", " ", t).strip()
    return ["X" if w == "POR" else w for w in t.split()]


def fon(w):
    """Clave fonética del español: b/v, s/z/c, k/q/c, g/j, ll/y, h muda, dobles."""
    w = re.sub(r"[^A-Z0-9]", "", w.upper())
    if w.isdigit():
        return w
    w = w.replace("CH", "#").replace("LL", "Y")
    w = re.sub(r"QU(?=[EI])", "K", w).replace("Q", "K")
    w = re.sub(r"GU(?=[EI])", "\x01", w)
    w = re.sub(r"G(?=[EI])", "J", w).replace("\x01", "G")
    w = re.sub(r"C(?=[EI])", "S", w).replace("Z", "S").replace("C", "K")
    w = w.replace("V", "B").replace("W", "U").replace("H", "").replace("X", "KS")
    return re.sub(r"(.)\1+", r"\1", w)


def _esqueleto(f):
    return "".join(c for c in f if c not in VOCALES)


def _dist(a, b):
    """Levenshtein."""
    if a == b:
        return 0
    if len(a) < len(b):
        a, b = b, a
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _subsec(corta, larga):
    it = iter(larga)
    return all(c in it for c in corta)


def similitud_token(q, t):
    """Parecido (0..1) entre una palabra dicha q y una palabra de la base t (ya en clave fonética)."""
    if q == t:
        return 1.0
    if q.isdigit() or t.isdigit():
        return 0.0
    lq, lt = len(q), len(t)
    mejor = 0.0
    if lt >= 2 and lq > lt and q.startswith(t):                      # CUAD <- CUADERNO
        mejor = 0.92 if lt >= 3 else 0.70
    elif lq >= 3 and lt > lq and t.startswith(q):                    # dijo menos que la base
        mejor = 0.78
    if lt >= 2 and lq >= lt + 2 and mejor < 0.85:                    # STD <- ESTANDAR
        l2 = q.lstrip("AEIOU") or q
        if (q[0] == t[0] or l2[0] == t[0]) and _subsec(t, l2):
            mejor = max(mejor, 0.78 if lt >= 3 else 0.66)
    if lq >= 3 and lt >= 3:
        e = 0.0
        if abs(lq - lt) <= max(2, 0.34 * max(lq, lt)):
            e = 1 - _dist(q, t) / max(lq, lt)
            if e >= 0.66:
                mejor = max(mejor, 0.95 * e)
        if e >= 0.45 and _esqueleto(q) == _esqueleto(t) and len(_esqueleto(q)) >= 3:
            mejor = max(mejor, 0.6 + 0.3 * e)
        if lt >= 4 and lq >= 4 and (t.endswith(q) or q.endswith(t)):
            mejor = max(mejor, 0.62)
    return mejor


def leer_alias(ruta):
    """alias.txt: 'hablado = ABREVIATURA' (una por línea, # comentarios)."""
    out = {}
    try:
        with open(ruta, encoding="utf-8-sig") as f:
            for ln in f:
                ln = ln.split("#")[0].strip()
                if "=" in ln:
                    a, b = ln.split("=", 1)
                    k = " ".join(tokens(a))
                    v = [x for x in tokens(b) if x]
                    if k and v:
                        out.setdefault(k, []).extend(v)
    except OSError:
        pass
    return out


class Buscador:
    def __init__(self, productos, alias=None):
        self.items, self.toks = [], []
        vistos = set()
        for p in productos:
            p = str(p).strip()
            if p and p not in vistos:
                vistos.add(p)
                self.items.append(p)
                self.toks.append(tokens(p))
        n = len(self.items)
        df, self.post = defaultdict(int), defaultdict(list)
        for i, tk in enumerate(self.toks):
            for t in set(tk):
                df[t] += 1
                self.post[t].append(i)
        self.vocab = list(df)
        self.idf = {t: math.log((n + 1) / (c + 1)) + 1.0 for t, c in df.items()}
        self.f = {t: fon(t) for t in self.vocab}
        self.fv = {}                    # pronunciaciones alternativas (Q + consonante = "cu")
        for t in self.vocab:
            if len(t) > 2 and t[0] == "Q" and t[1] not in VOCALES and t.isalpha():
                self.fv[t] = "KU" + fon(t[1:])
        self.peso_t = {}
        for t in self.vocab:                       # códigos y letras sueltas casi no se dictan
            w = self.idf[t]
            if t.isdigit() and len(t) >= 4:
                w *= 0.3
            elif len(t) == 1 and not t.isdigit():
                w *= 0.5
            self.peso_t[t] = w
        self.ptot = [sum(self.peso_t[t] for t in set(tk)) or 1.0 for tk in self.toks]
        self.alias = {k: list(v) for k, v in ALIAS_BASE.items()}
        for k, v in (alias or {}).items():
            self.alias.setdefault(k, []).extend(v)
        self._cache = {}
        self.df = {t: len(self.post[t]) for t in self.vocab}
        self._tri, self._esq, self._cortas = defaultdict(set), defaultdict(set), []
        for t in self.vocab:
            if t.isdigit() or not self.f[t]:
                continue
            ft = self.f[t]
            for v in (ft, self.fv.get(t)):
                if v:
                    for g in self._trigramas(v):
                        self._tri[g].add(t)
                    self._esq[_esqueleto(v)].add(t)
            if 2 <= len(ft) <= 5:
                self._cortas.append(t)

    @staticmethod
    def _trigramas(f):
        p = " " + f + " "
        return {p[i:i + 3] for i in range(len(p) - 2)}

    # ---- parecidos de una palabra dicha con todo el vocabulario ----
    def _sims(self, w):
        if w in self._cache:
            return self._cache[w]
        out = {}
        if w.isdigit():
            if w in self.idf:
                out[w] = 1.0
        else:
            fq = fon(w)
            if fq:
                cand = set()
                for g in self._trigramas(fq):
                    cand |= self._tri.get(g, set())
                cand |= self._esq.get(_esqueleto(fq), set())
                if len(fq) >= 4:
                    cand.update(self._cortas)
                for t in cand:
                    s = similitud_token(fq, self.f[t])
                    if t in self.fv:
                        s = max(s, similitud_token(fq, self.fv[t]))
                    if s >= 0.6:
                        out[t] = s
            for t in self.alias.get(w, []):
                if t in self.idf:
                    out[t] = max(out.get(t, 0), 0.95)
        if len(self._cache) > 4000:
            self._cache.clear()
        self._cache[w] = out
        return out

    def _analizar(self, frase):
        q = tokens(frase)
        sims = [self._sims(w) for w in q]
        joins = {}
        for i in range(len(q) - 1):
            a, b = q[i], q[i + 1]
            if a.isalpha() and b.isalpha() and len(a + b) >= 4:
                # la unión solo vale si explica MEJOR que las dos palabras por separado
                s = {t: v for t, v in self._sims(a + b).items()
                     if v >= 0.72 and v >= max(sims[i].get(t, 0), sims[i + 1].get(t, 0)) + 0.08}
                if s:
                    joins[i] = s
        return q, sims, joins

    def buscar(self, frase, k=5):
        """[(score 0..1, producto, parte_del_nombre_dicha 0..1)]"""
        q, sims, joins = self._analizar(frase)
        if not q:
            return []
        n = len(q)
        # peso de cada palabra dicha = rareza de la palabra de la base que mejor le corresponde
        pesos = []
        for i in range(n):
            cand = list(sims[i].items())
            for j in (i - 1, i):
                if j in joins:
                    cand += list(joins[j].items())
            if cand:
                t, s = max(cand, key=lambda x: (x[1], self.idf[x[0]]))
                w = self.peso_t[t]
            else:
                w = 2.0
            if q[i] in RELLENO:
                w *= 0.3
            pesos.append(max(w, 0.3))
        total = sum(pesos)
        mejor = [defaultdict(float) for _ in range(n)]
        cubierto = defaultdict(set)
        for i in range(n):
            fuentes = [(sims[i], i)]
            for j in (i - 1, i):
                if j in joins:
                    fuentes.append((joins[j], i))
            for d, _ in fuentes:
                for t, s in d.items():
                    if s < 0.6 or len(self.post[t]) > 1400:
                        continue
                    for p in self.post[t]:
                        if s > mejor[i][p]:
                            mejor[i][p] = s
                        if s >= 0.75:
                            cubierto[p].add(t)
        cands = set()
        for m in mejor:
            cands.update(m)
        nums = [w for w in q if w.isdigit() and len(w) <= 4]
        res = []
        for p in cands:
            qc = sum(pesos[i] * mejor[i].get(p, 0.0) for i in range(n)) / total
            pc = sum(self.peso_t[t] for t in cubierto.get(p, ())) / self.ptot[p]
            sc = 0.62 * qc + 0.38 * min(pc, 1.0)
            if nums:
                tk = set(self.toks[p])
                sc -= min(0.2, 0.08 * sum(1 for w in nums if w not in tk))
            res.append((sc, p, pc))
        res.sort(key=lambda x: -x[0])
        return [(round(min(max(sc, 0.0), 1.0), 3), self.items[p], round(min(pc, 1.0), 3))
                for sc, p, pc in res[:k]]

    def _elegir(self, d):
        """De los parecidos, el mejor; si hay empate técnico, el que más aparece en la base."""
        top = max(d.values())
        t = max((t for t, s in d.items() if s >= top - 0.10),
                key=lambda t: d[t] * (1 + 0.15 * math.log10(self.df[t] + 1)))
        return t, d[t]

    def corregir(self, frase):
        """Lo que dijiste, con las palabras escritas como están en la base
        ('cuaderno kinota' -> 'CUAD QNOTA'). Lo no reconocido queda como se oyó."""
        q, sims, joins = self._analizar(frase)
        out, i = [], 0
        while i < len(q):
            w = q[i]
            sj = joins.get(i)
            mj = self._elegir(sj) if sj else None
            ms = self._elegir(sims[i]) if sims[i] else None
            if mj and mj[1] >= 0.74 and (not ms or mj[1] > ms[1] + 0.05) and i + 1 < len(q) \
                    and not (sims[i + 1] and max(sims[i + 1].values()) >= 0.95
                             and sims[i] and max(sims[i].values()) >= 0.95):
                out.append(mj[0])
                i += 2
                continue
            if w in RELLENO and not (ms and ms[1] >= 1.0):
                i += 1
                continue
            if w.isdigit():
                out.append(w)
            elif ms and ms[1] >= 0.75:
                out.append(ms[0])
            else:
                out.append(w)
            i += 1
        return " ".join(out)


# --------------------------------------------------------------------
SENSIBILIDAD = {      # nombre: (similitud mínima, ventaja sobre el 2.º, cuánto del nombre debes decir)
    "estricta": (0.80, 0.05, 0.75),
    "normal": (0.70, 0.04, 0.55),
    "flexible": (0.55, 0.00, 0.00),
}


def decidir_coincidencia(buscador, frase, params):
    """-> None | dict(modo='auto'|'ambiguo'|'nada', texto, score, alts, corregido)"""
    if buscador is None:
        return None
    umbral, margen, cobertura = params
    r = buscador.buscar(frase, k=5)
    corr = buscador.corregir(frase)
    if not r:
        return {"modo": "nada", "texto": "", "score": 0.0, "alts": [], "corregido": corr}
    top, cub = r[0][0], r[0][2]
    segundo = r[1][0] if len(r) > 1 else 0.0
    alts = [x[1] for x in r[:5]]
    ventaja = top - segundo
    claro = top >= 0.97 or (ventaja >= margen and (cub >= cobertura or ventaja >= 0.10))
    if top >= umbral and claro:
        return {"modo": "auto", "texto": r[0][1], "score": top, "alts": alts, "corregido": corr}
    if top >= max(0.50, umbral - 0.20):
        return {"modo": "ambiguo", "texto": "", "score": top, "alts": alts, "corregido": corr}
    return {"modo": "nada", "texto": "", "score": top, "alts": [], "corregido": corr}
