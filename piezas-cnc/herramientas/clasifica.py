# -*- coding: utf-8 -*-
"""Clasifica cada pieza por material a partir del texto que va detras del codigo."""
import json,re,collections,unicodedata

RX_COD=re.compile(r'(?<![0-9])(\d{6,9})-(\d{3,4})-(\d{2,4})(?:[_-]([A-Za-z]))?(?![0-9A-Za-z])')

def sin_tildes(s):
    return ''.join(c for c in unicodedata.normalize('NFD',s) if unicodedata.category(c)!='Mn')

def limpia(nombre):
    n=re.sub(r'\.ard$','',nombre,flags=re.I)
    m=RX_COD.search(n)
    t=n[m.end():] if m else n
    t=re.sub(r'[\(\[\{][^\)\]\}]*[\)\]\}]',' ',t)
    t=re.sub(r'cara\s*(buena|mala)[^_]*',' ',t,flags=re.I)
    t=re.sub(r'\bcara\s*[_\-\.]?\s*[ab]\b',' ',t,flags=re.I)
    t=re.sub(r'[_\s]\d+\s*(und|uds|ud|u)\b\.?',' ',t,flags=re.I)
    t=re.sub(r'\b\d+\s*(und|uds|ud)\b\.?',' ',t,flags=re.I)
    t=re.sub(r'[_\-]+',' ',t)
    return re.sub(r'\s+',' ',t).strip(' .,-_')

FAMILIAS=[
 ('PRELACADO IGNIFUGO', r'(prelacad|prelac|prela|prlacad)\w*\s*igni|igni\w*\s*prelac'),
 ('PRELACADO',          r'prelacad|prelac|prela\d|\bprela\b|prlacad'),
 ('LACADO',             r'lacad|\b(gris|blanco|negro|blanca|negra)\s*(brillo|mate)\b'),
 ('DM / MDF IGNIFUGO',  r'\b(dm|mdf)\w*[^0-9]{0,14}igni|igni[^0-9]{0,14}\b(dm|mdf)'),
 ('DM / MDF HIDROFUGO', r'hidrofug|hidrofob'),
 ('DM / MDF NEGRO',     r'\b(dm|mdf)\w*[^0-9]{0,14}negr'),
 ('VALCHROMAT',         r'valchro'),
 ('DM / MDF',           r'\bdm|\bmdf|\bdmf'),
 ('AGLOMERADO / MELAMINA', r'aglo|melamin|\bmel\b|\bmelam'),
 ('CONTRACHAPADO',      r'contrachapad|plywood'),
 ('OSB',                r'\bosb'),
 ('CHOPO',              r'chopo'),
 ('ALPI',               r'\balpi'),
 ('ARCE / MAPLE',       r'\barce|maple'),
 ('ROBLE',              r'roble'),
 ('NOGAL',              r'nogal'),
 ('FRESNO',             r'fresno'),
 ('HAYA',               r'\bhaya'),
 ('PINO / ALISTONADO',  r'aliston|\bpino'),
 ('UNDORA CAMPASPERO',  r'campaspero'),
 ('UNDORA',             r'undora'),
 ('KERABAN',            r'keraban'),
 ('TARKETT',            r'tarkett'),
 ('TABLERO CHINO',      r'chino'),
 ('KRION',              r'krion'),
 ('CORIAN',             r'corian'),
 ('COMPAC / FENIX',     r'compac|fenix'),
 ('HPL',                r'\bhpl'),
 ('ECOPANEL',           r'ecopanel'),
 ('PANESPOL',           r'panespol'),
 ('PVC ESPUMADO',       r'\bpvc|espumad|foamx|forex'),
 ('METACRILATO',        r'metacril|\bmt\b|\bmt[\s\-]?xt|\bmtxt|\bmt\d|\bmeta\b|\bopal'),
 ('POLICARBONATO',      r'policarbon'),
 ('PET',                r'\bpet\b|\bpet\d'),
 ('DIBOND / COMPOSITE', r'dibond|composite|material compuesto'),
 ('CARTON / PANAL',     r'carton|\bpanal'),
 ('ALUMINIO',           r'aluminio'),
 ('POLIESTIRENO / XPS', r'poliestiren|\bxps\b|porex'),
 ('RECHAPADO',          r'rechapad|\brech\b|chapad'),
 ('MADERA GENERICO',    r'generico'),
 ('KERABAN',            r'karaban'),
 ('FIBRAPAN',           r'fibrapan'),
 ('FLEXIBLE',           r'flexible'),
 ('LACADO',             r'\bbrillo\b'),
 ('METACRILATO',        r'incoloro|\bmtinc'),
]
FAM_RX=[(c,re.compile(r,re.I)) for c,r in FAMILIAS]
RX_GRUESO=re.compile(r'(?<![\d.,])(\d{1,3}(?:[.,]\d{1,2})?)\s*(?:mm)?(?![\d.,%])',re.I)

def grueso(t):
    for m in RX_GRUESO.finditer(t):
        try: f=float(m.group(1).replace(',','.'))
        except ValueError: continue
        if 0.5<=f<=120: return '%g'%f
    return None

def clasifica(nombre):
    t=limpia(nombre)
    if not t: return None,None,''
    s=sin_tildes(t).lower()
    for canon,rx in FAM_RX:
        if rx.search(s): return canon,grueso(s),t
    return None,grueso(s),t

if __name__=='__main__':
    win=json.load(open('piezas_a_copiar.json',encoding='utf-8'))
    sinfam=collections.Counter(); g=collections.Counter(); vacio=0
    for w in win:
        f,e,t=clasifica(w['nombre'])
        if not t: vacio+=1
        elif not f: sinfam[t]+=1
        else: g[(f,e)]+=1
    print('piezas:',len(win),'| sin texto tras el codigo:',vacio)
    print('clasificadas:',sum(g.values()),'(%.1f%%)'%(100*sum(g.values())/len(win)),
          '| sin familia:',sum(sinfam.values()),'(%d textos distintos)'%len(sinfam))
    print('grupos material+espesor:',len(g))
    print('--- top 25 sin clasificar ---')
    for k,v in sinfam.most_common(25): print('%6d  %r'%(v,k))
    fam=collections.Counter()
    for (f,e),v in g.items(): fam[f]+=v
    print('--- familias (%d) ---'%len(fam))
    for k,v in fam.most_common(): print('%6d  %s'%(v,k))
