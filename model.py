import json, io
import numpy as np
import openpyxl
D = {}
EXPECT = {'basic': 380, 'small': 165, 'mid': 83, 'large': 33}

def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)

def initialize(level, raw):
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    base = book['국산투입계수(Ad)']
    codes = [str(c) for c in base[4][2:] if c is not None and not str(c).startswith('9')]
    if len(codes) != EXPECT[level] or len(set(codes)) != len(codes):
        raise ValueError('%s 파일의 부문 수가 %d개가 아닙니다(읽은 수: %d). 올바른 파일인지 확인하세요.' % (level, EXPECT[level], len(codes)))
    nm = {str(r[0]): str(r[1]) for r in base[6:] if r and r[0] is not None}
    names = [nm[c] for c in codes]
    def matrix(s):
        rows = book[s]
        ri = {str(r[0]): i for i, r in enumerate(rows) if r and r[0] is not None}
        ci = {str(c): i for i, c in enumerate(rows[4])}
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all(): raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a
    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)'); am = matrix('수입투입계수(Am)')
    rows = book['총투입계수(A)']; ci = {str(c): i for i, c in enumerate(rows[4])}
    vr = next(r for r in rows if r and str(r[0]) == '9690')
    v = np.array([vr[ci[c]] for c in codes], dtype=float)
    I = np.eye(len(codes))
    residuals = {'생산': float(np.max(np.abs((I - ad) @ L - I))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    if max(residuals.values()) >= 1e-8: raise ValueError('계수 검증 실패: ' + str(residuals))
    D[level] = dict(book=book, codes=codes, names=names, L=L, M=M, V=V)
    return emit({'sectors': [{'code': c, 'name': n} for c, n in zip(codes, names)],
                 'sheets': list(book), 'residuals': residuals})

def load_class(raw):
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    rows = list(wb[wb.sheetnames[0]].values); wb.close()
    cur = [None, None, None]; out = {}
    for r in rows[3:]:
        if r[0] is None or not str(r[0]).isdigit(): continue
        for k in range(3):
            if r[2 + 2 * k] is not None: cur[k] = r[3 + 2 * k]
        out[str(r[0])] = ' > '.join(str(x).strip() for x in (cur[2], cur[1], cur[0]) if x)
    return emit(out)

def explore(level, sheet, rowcode):
    rows = D[level]['book'][sheet]
    choices = [{'code': str(r[0]), 'name': str(r[1])} for r in rows[6:] if len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if str(r[0]) == rowcode), None)
    if chosen is None: chosen = next(r for r in rows[6:] if r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None: continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)): val = float(val)
        elif val is not None: val = str(val)
        cells.append({'code': str(rows[4][i]), 'name': str(rows[5][i] or ''), 'value': val})
    return emit({'rows': choices, 'selected': str(chosen[0]), 'name': str(chosen[1]),
                 'unit': str(rows[3][0] or '係数').replace('係数', '계수 / 비율: 행·열 명칭 확인'), 'cells': cells})

def simulate(level, payload):
    d = D[level]; codes = d['codes']
    f = np.zeros(len(codes))
    for s in json.loads(payload):
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0: raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not f.sum() > 0: raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')
    x = d['L'] @ f; m = d['M'] @ f; v = d['V'] @ f
    result = [{'code': c, 'name': n, 'production': float(a), 'imports': float(b), 'valueAdded': float(e)}
              for c, n, a, b, e in zip(codes, d['names'], x, m, v)]
    result.sort(key=lambda r: r['production'], reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
                 'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result})
