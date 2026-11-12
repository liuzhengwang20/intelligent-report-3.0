# -*- coding: utf-8 -*-
"""
财务报表智能填报工具 3.0
支持两种会计准则：
  1. 一般企业准则：上传 UI1(当期excel) + UI2(上年年报或第四季度pdf) + UI3(上年X季度pdf)
  2. 小企业会计准则：上传 UI1(当期excel) + UI2(上年年报或第四季度pdf) + UI4(小企业上季度财报pdf)
小企业利润表：本年累计金额 ← 当期excel 本期发生额；
  本期金额：第1季度(UI1与UI4年份不同)=本年累计金额；
            第2~4季度(同年)=当期本年累计-上期本年累计(UI4 pdf的"本年累计金额"列)。
信用减值损失（小企业准则）：当期excel含该科目时自动重分类——
  本年累计金额=|信用减值损失累计|(如excel另有营业外支出则相加)；
  本期金额=重分类后本年累计-上期pdf营业外支出本年累计(即本季度新增信用减值损失)；
  营业利润/利润总额/净利润改按模板上行科目勾稽计算(不做两表直接相减)；
  并对小企业利润表做勾稽校验，结果写入校验日志"汇总"页。

3.0 修复内容：
  1. 复制模板时保留原调色板（修复输出单元格颜色变暗发黑的问题）
  2. PDF解析时排除"行次"列数字（修复数据前面多出1、2、3...行次号的问题）
  3. UI4 改为上传pdf文件，取其上季度利润表"本年累计金额"列参与运算
仅填报数据，不填充单元格格式，不修改源文件和模板。
"""
import os
import re
import sys
import json
import shutil
import datetime
import unicodedata
import traceback

APP_DIR = os.path.dirname(os.path.abspath(
    sys.executable if getattr(sys, 'frozen', False) else __file__))
TPL_DIR = os.path.join(APP_DIR, '模板')
OUT_DIR = os.path.join(APP_DIR, '填报结果')
MAP_PATH = os.path.join(APP_DIR, '科目映射.json')

# 负债/权益分界：模板行次小于该值算负债部分
LIAB_THRESHOLD = {'general': 61, 'small': 48}
# N04 行次阈值：负债部分行次 < 90
N04_LIAB_MAX_ROWNO = 90
# 行次列数字的排除半径（pt）：数值词中心离"行次"表头该距离内的视为行号
ROWNO_EXCLUDE_R = 30

DEFAULT_ALIASES = {
    '应收账款': ['应收账款净额'],
    '其他应收款': ['其他应收款净额'],
    '存货': ['存货净额'],
    '合同资产': ['合同资产净额'],
    '长期股权投资': ['长期股权投资净额'],
    '投资性房地产': ['投资性房地产净额'],
    '固定资产': ['固定资产净额'],
    '在建工程': ['在建工程净额'],
    '使用权资产': ['使用权资产净值'],
    '无形资产': ['无形资产净额'],
    '商誉': ['商誉净额'],
    '应收款项融资': ['应收账款融资'],
    '持有待售负债': ['划分为持有待售的负债'],
    '实收资本（或股本）': ['股本（或实收资本）'],
    '长期应付款': ['长期应付款净额'],
    '（一）基本每股收益': ['（一）基本每股收益（元/股）'],
    '（二）稀释每股收益': ['（二）稀释每股收益（元/股）'],
    # 小企业模板专用
    '预付账款': ['预付款项'],
    '预收账款': ['预收款项'],
    '固定资产原价': ['固定资产'],
    '固定资产账面价值': ['固定资产净额'],
    '资产合计': ['资产总计'],
    '其中：利息费用（收入以“-”号填列）': ['其中：利息费用'],
}


def load_aliases():
    raw = {}
    if os.path.exists(MAP_PATH):
        try:
            with open(MAP_PATH, encoding='utf-8') as f:
                m = json.load(f).get('aliases', {})
            if m:
                raw = m
        except Exception:
            pass
    if not raw:
        raw = DEFAULT_ALIASES
    return {norm_key(k): [norm_key(v) for v in vals]
            for k, vals in raw.items()}


def norm_key(s):
    """科目名标准化：全半角统一、去空白、去装饰点、其它→其他"""
    if s is None:
        return ''
    s = str(s)
    s = unicodedata.normalize('NFKC', s)
    s = re.sub(r'\s+', '', s)
    s = s.replace('其它', '其他')
    s = s.replace('．', '').replace('.', '')
    s = re.sub(r'[–—−]', '-', s)
    return s


def parse_num(raw):
    """解析数值：支持千分位、括号负数、负号"""
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip().replace(',', '').replace('，', '')
    if not s:
        return None
    m = re.match(r'^（?\s*-?\s*([\d.]+)\s*）?$', s)
    if m:
        try:
            val = float(m.group(1))
        except ValueError:
            return None  # 类似 '359354.74359354.74' 的异常值
        if s.strip().startswith('-') or s.strip().startswith('（'):
            val = -val
        return val
    return None


def extract_year(fname):
    """从文件名提取第一个20xx年份，供季度判断"""
    m = re.search(r'20\d{2}', os.path.basename(fname))
    return int(m.group(0)) if m else None


# ---------------------------------------------------------------- 模板读取

def get_template_path(kind):
    """返回对应准则的模板xls路径；kind='general'或'small'"""
    if not os.path.isdir(TPL_DIR):
        raise FileNotFoundError('模板文件夹不存在，请先通过"替换模板"放入模板')
    files = [f for f in os.listdir(TPL_DIR) if f.lower().endswith('.xls')]
    if not files:
        raise FileNotFoundError('模板文件夹中没有xls模板文件，请先通过"替换模板"放入模板')
    for f in files:
        is_small = '小企业' in f
        if kind == 'small' and is_small:
            return os.path.join(TPL_DIR, f)
        if kind == 'general' and not is_small:
            return os.path.join(TPL_DIR, f)
    raise FileNotFoundError(
        '模板文件夹中缺少%s模板' % ('小企业会计准则' if kind == 'small' else '一般企业'))


def read_template_spec(kind):
    """读取模板xls，返回填报点列表
    [{sheet, row0, col0, colname, item, side, fill_type}]
    fill_type: current(当期excel) / annual(UI2 pdf) / quarter_prev(UI3 pdf) / small_calc(计算)"""
    import xlrd
    tp = get_template_path(kind)
    wb = xlrd.open_workbook(tp)
    spec = []
    thr = LIAB_THRESHOLD[kind]
    ws = wb.sheet_by_name('资产负债表')

    def is_rownum(v):
        try:
            float(v)
            return str(v).strip() not in ('', '行次')
        except (TypeError, ValueError):
            return False

    bs_colname_prev = '上年年末余额' if kind == 'general' else '年初余额'
    for r in range(ws.nrows):
        item_a = str(ws.cell_value(r, 1)).strip()
        rowno_a = ws.cell_value(r, 2)
        if item_a and is_rownum(rowno_a):
            spec.append({'sheet': '资产负债表', 'row0': r, 'col0': 3,
                         'colname': '期末余额', 'item': item_a,
                         'side': 'asset', 'fill_type': 'current'})
            spec.append({'sheet': '资产负债表', 'row0': r, 'col0': 4,
                         'colname': bs_colname_prev, 'item': item_a,
                         'side': 'asset', 'fill_type': 'annual'})
        item_l = str(ws.cell_value(r, 5)).strip()
        rowno_l = ws.cell_value(r, 6)
        if item_l and is_rownum(rowno_l):
            side = 'liab' if float(rowno_l) < thr else 'equity'
            spec.append({'sheet': '资产负债表', 'row0': r, 'col0': 7,
                         'colname': '期末余额', 'item': item_l,
                         'side': side, 'fill_type': 'current'})
            spec.append({'sheet': '资产负债表', 'row0': r, 'col0': 8,
                         'colname': bs_colname_prev, 'item': item_l,
                         'side': side, 'fill_type': 'annual'})
    # 利润表：sheet名因准则而异
    pl_sheet = '利润表' if '利润表' in wb.sheet_names() else '利润表_月季报'
    ws = wb.sheet_by_name(pl_sheet)
    hdr_r = None
    for r in range(ws.nrows):
        if str(ws.cell_value(r, 1)).strip() == '项目' and \
                str(ws.cell_value(r, 2)).strip() in ('行次', '本期金额'):
            hdr_r = r
            break
    if hdr_r is not None:
        for r in range(hdr_r + 1, ws.nrows):
            item = str(ws.cell_value(r, 1)).strip()
            if not item:
                continue
            if kind == 'general':
                spec.append({'sheet': pl_sheet, 'row0': r, 'col0': 2,
                             'colname': '本期金额', 'item': item,
                             'side': 'pl', 'fill_type': 'current'})
                spec.append({'sheet': pl_sheet, 'row0': r, 'col0': 3,
                             'colname': '上期金额', 'item': item,
                             'side': 'pl', 'fill_type': 'quarter_prev'})
            else:
                spec.append({'sheet': pl_sheet, 'row0': r, 'col0': 4,
                             'colname': '本年累计金额', 'item': item,
                             'side': 'pl', 'fill_type': 'current'})
                spec.append({'sheet': pl_sheet, 'row0': r, 'col0': 3,
                             'colname': '本期金额', 'item': item,
                             'side': 'pl', 'fill_type': 'small_calc'})
    return spec


def small_pl_equations(tpath):
    """解析小企业模板利润表的勾稽结构。
    顶层科目行无缩进；"其中"明细行有缩进/含"其中"字样，不参与勾稽。
    返回 {'sheet', 'totals': {rev/op/total/net: 行号},
          'segA'(收入→营业利润), 'segB'(营业利润→利润总额), 'segC'(利润总额→净利润)}
    每段元素为 (行号, 符号, 科目名)：减项-1、加项+1、减值损失行+1(以负数列示)"""
    import xlrd
    wb = xlrd.open_workbook(tpath)
    name = '利润表' if '利润表' in wb.sheet_names() else '利润表_月季报'
    ws = wb.sheet_by_name(name)
    rows = []
    for r in range(ws.nrows):
        raw = str(ws.cell_value(r, 1))
        s = raw.strip()
        if not s or s == '项目':
            continue
        indent = len(raw) - len(raw.lstrip())
        rows.append({'r': r, 'item': s, 'nk': norm_key(s), 'indent': indent})
    totals = {}
    for d in rows:
        for key, pfx in (('rev', '一、营业收入'), ('op', '二、营业利润'),
                         ('total', '三、利润总额'), ('net', '四、净利润')):
            if d['nk'].startswith(norm_key(pfx)):
                totals[key] = d['r']
    if any(k not in totals for k in ('rev', 'op', 'total', 'net')):
        return None
    sub_hint = ('其中', '消费税', '城市维护建设税', '资源税', '土地增值税',
                '城镇土地使用税', '教育费附加', '商品维修费', '广告费和业务宣传费',
                '开办费', '业务招待费', '研究费用', '利息费用', '利息收入',
                '政府补助', '坏账损失', '无法收回的长期债券投资损失',
                '无法收回的长期股权投资损失', '自然灾害', '税收滞纳金')

    def is_sub(d):
        return d['indent'] > 0 or '其中' in d['item'] or \
            any(h in d['item'] for h in sub_hint)

    def seg(a, b):
        out = []
        for d in rows:
            if totals[a] < d['r'] < totals[b] and not is_sub(d):
                sign = 1.0 if (d['nk'].startswith('加:') or '减值损失' in d['nk']) \
                    else -1.0
                out.append((d['r'], sign, d['item']))
        return out

    return {'sheet': name, 'totals': totals,
            'segA': seg('rev', 'op'), 'segB': seg('op', 'total'),
            'segC': seg('total', 'net')}


# ---------------------------------------------------------------- N04 excel 读取

def read_n04(path):
    """读取N04当期excel → {section: {normkey: {val, raw, name, begin}}}
    列位置按表头名称定位，兼容一般/小企业不同版式"""
    import openpyxl
    data = {'asset': {}, 'liab': {}, 'equity': {}, 'pl': {}}
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)

    def cell(row, i):
        return row[i] if row and i < len(row) else None

    def nz(v):
        return v if v is not None and str(v).strip() != '' else None

    ws = wb['资产负债表']
    rows = list(ws.iter_rows(values_only=True))
    hdr = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip() == '资产' and str(row[1]).strip() == '行次':
            hdr = i
            break
    if hdr is not None:
        h = [str(c).strip() if c is not None else '' for c in rows[hdr]]
        # 资产侧列号
        a_end = [j for j in range(1, len(h)) if str(h[j]).startswith('负债')]
        a_end = a_end[0] if a_end else 5
        a_q = [j for j in range(1, a_end) if h[j] == '期末余额']
        a_b = [j for j in range(1, a_end) if h[j] in ('期初余额', '年初余额')]
        a_q = a_q[0] if a_q else 2
        a_b = a_b[0] if a_b else 3
        # 负债和所有者权益侧列号
        l_q = [j for j in range(a_end, len(h)) if h[j] == '期末余额']
        l_b = [j for j in range(a_end, len(h)) if h[j] in ('期初余额', '年初余额')]
        l_q = l_q[0] if l_q else 6
        l_b = l_b[0] if l_b else 7
        for row in rows[hdr + 1:]:
            if not row:
                continue
            name_a = str(cell(row, 0)).strip() if cell(row, 0) else ''
            rowno_a = cell(row, 1)
            if name_a and rowno_a is not None and str(rowno_a).strip() != '':
                data['asset'][norm_key(name_a)] = {
                    'val': nz(cell(row, a_q)), 'raw': cell(row, a_q),
                    'name': name_a, 'begin': nz(cell(row, a_b))}
            name_l = str(cell(row, a_end)).strip() if cell(row, a_end) else ''
            rowno_l = cell(row, a_end + 1)
            if name_l and rowno_l is not None and str(rowno_l).strip() != '':
                sec = 'liab' if float(rowno_l) < N04_LIAB_MAX_ROWNO else 'equity'
                data[sec][norm_key(name_l)] = {
                    'val': nz(cell(row, l_q)), 'raw': cell(row, l_q),
                    'name': name_l, 'begin': nz(cell(row, l_b))}
    ws = wb['利润表']
    rows = list(ws.iter_rows(values_only=True))
    hdr = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip() == '项目' and str(row[1]).strip() == '行次':
            hdr = i
            break
    if hdr is not None:
        h = [str(c).strip() if c is not None else '' for c in rows[hdr]]
        # 取"本期发生额"列，找不到再取"本年累计金额"/"本月发生额"
        col = None
        for cand in ('本期发生额', '本年累计金额', '本月发生额', '本期金额'):
            if cand in h:
                col = h.index(cand)
                break
        if col is None:
            col = 3
        for row in rows[hdr + 1:]:
            if not row:
                continue
            name = str(cell(row, 0)).strip() if cell(row, 0) else ''
            if not name:
                continue
            val = cell(row, col)
            data['pl'][norm_key(name)] = {
                'val': nz(val), 'raw': val, 'name': name}
    wb.close()
    return data


# ---------------------------------------------------------------- PDF 读取

NUM_TK = re.compile(r'^-?[\d,，]+(?:\.\d+)?$|^（-?[\d,，]+(?:\.\d+)?）$|^\(-?[\d,，]+(?:\.\d+)?\)$')


def _split_mixed(tok):
    """拆分数字与文字粘连的词，如 '51,408,589.78流动负债合计'"""
    parts = re.findall(r'-?[\d,，]+(?:\.\d+)?|（\)）?|[^\d]+', tok)
    return [p for p in parts if p]


def _rows_from_page(page):
    """按行聚合词，返回 [{y, tokens:[(x0, x1, text)]}]，并拆分粘连词"""
    rows = {}
    for w in page.get_text('words'):
        x0, y0, x1, txt = w[0], w[1], w[2], w[4]
        toks = _split_mixed(txt)
        if len(toks) == 1:
            rows.setdefault(round(y0 / 5), []).append((x0, x1, txt, y0))
        else:
            total = sum(len(t) for t in toks)
            wdt = max(x1 - x0, 1)
            cur = x0
            for t in toks:
                w2 = wdt * len(t) / total
                rows.setdefault(round(y0 / 5), []).append((cur, cur + w2, t, y0))
                cur += w2
    out = []
    for k, toks in rows.items():
        toks.sort(key=lambda t: t[0])
        out.append({'y': min(t[3] for t in toks),
                    'tokens': [(x0, x1, t) for x0, x1, t, _ in toks]})
    out.sort(key=lambda d: d['y'])
    return out


def _is_rownum_tok(x0, x1, rowno_xs):
    """数值词的中心是否落在"行次"列附近 → 是行号，应排除"""
    if not rowno_xs:
        return False
    cx = (x0 + x1) / 2
    return any(abs(cx - x) < ROWNO_EXCLUDE_R for x in rowno_xs)


def _split_merged_rownum(x0, x1, t, rowno_xs, rowno_rights=None):
    """处理"行次号与数值粘连"的词（如 '33'与'44,803,412.23'被提取成
    '3344,803,412.23'）：仅当词明显侵入行次列（起点比行次右边界更靠左
    1.5个字符宽以上）才拆分，行号占1~2个字符（数字等宽）。
    返回 (新x0, 数值部分)；不需要拆分时返回 None"""
    if not rowno_xs or len(t) <= 5:
        return None
    for rx in rowno_xs:
        row_right = (rowno_rights or {}).get(rx, rx + 12)
        cw = max((x1 - x0) / len(t), 2.5)  # 数字等宽估计
        if x0 < row_right - cw * 1.5 and x1 > row_right + 8:
            idx = (row_right - x0) / cw
            k = 1 if idx < 1.5 else 2
            if k < len(t) - 1 and t[:k].isdigit() and NUM_TK.match(t[k:]):
                return x0 + cw * k, t[k:]
    return None


def _measure_rownum_rights(rows, start, rowno_xs):
    """用被排除的行号词实测各侧行次列的右边界（行号右对齐）"""
    rights = {}
    if not rowno_xs:
        return rights
    for d in rows[start:]:
        for x, x1, t in d['tokens']:
            if NUM_TK.match(t.replace(' ', '')) and _is_rownum_tok(x, x1, rowno_xs):
                cx = (x + x1) / 2
                nearest = min(rowno_xs, key=lambda rx: abs(cx - rx))
                rights[nearest] = max(rights.get(nearest, 0), x1)
    return rights


BS_COL_LABELS = ('期末余额', '年初余额', '上年年末余额', '期末', '上年年末', '年初')
PL_COL_LABELS = ('本期金额', '上期金额', '本年累计金额', '本月金额', '上年金额',
                 '本月发生额', '本年累计')


def _bs_header(rows):
    """资产负债表页：返回(4个数值列中心, 数据起始行号, 行次列x列表)
    支持 期末余额/年初余额 整词表头 或 期末/余额 分列表头"""
    hdr_i = None
    for i, d in enumerate(rows):
        hits = [t for _, _, t in d['tokens'] if t in BS_COL_LABELS]
        if len(hits) >= 4:
            hdr_i = i
            break
    if hdr_i is None:
        return [220.75, 275.05, 484.65, 538.9], 0, []
    xs = sorted(x for x, _, t in rows[hdr_i]['tokens'] if t in BS_COL_LABELS)
    rowno_xs = [x for x, _, t in rows[hdr_i]['tokens'] if t == '行次']
    if len(xs) < 4:
        return [220.75, 275.05, 484.65, 538.9], hdr_i + 1, rowno_xs
    centers = [x + 20 for x in xs]
    start = hdr_i + 1
    for j in range(hdr_i + 1, min(hdr_i + 4, len(rows))):
        toks = [t for _, _, t in rows[j]['tokens']]
        if len([t for t in toks if t == '余额']) >= 2 and not any(
                t in BS_COL_LABELS for t in toks):
            start = j + 1
    return centers, start, rowno_xs


def _pl_header(rows):
    """利润表页：返回(列中心及含义列表[(x, 'cur'|'cum')], 数据起始行号, 行次列x列表)
    含义：含"累计"字样 → cum（本年累计列），其余 → cur（本期/本月列）"""
    for i, d in enumerate(rows):
        toks = [t for _, _, t in d['tokens']]
        if '项目' in toks and any(t in PL_COL_LABELS for t in toks):
            rowno_xs = []
            cols = []
            for x, _, t in d['tokens']:
                if t == '行次':
                    rowno_xs.append(x)
                elif t in PL_COL_LABELS:
                    kind = 'cum' if '累计' in t else 'cur'
                    cols.append((x + 20, kind))
            if cols:
                cols.sort(key=lambda c: c[0])
                return cols, i + 1, rowno_xs
    return [(361, 'cur'), (500, 'cur')], 0, []


def _assign_by_center(vals, centers):
    """把值词按词中心(x0+x1)/2与最近列中心的距离归类"""
    out = {}
    for x, x1, t in vals:
        cx = (x + x1) / 2
        best = min(centers, key=lambda c: abs(cx - c))
        out.setdefault(best, []).append(t)
    return out


# 同一单元格内换行时，两行文字的行距明显小于正常行距（约15pt）
NAME_LINE_GAP = 10.0


def _is_section_header(nk):
    """是否为"所有者权益（或股东权益）："这类分段标题（不是科目行）"""
    return (nk.startswith('所有者权益') and nk.endswith(('：', ':'))
            and '合计' not in nk and '总计' not in nk)


def _not_item_name(nk):
    """段标题、省略号等不是科目名的内容"""
    return (not nk) or nk.endswith(('：', ':')) or '……' in nk


def _parse_bs_pdf_page(rows):
    """解析资产负债表页 → {'asset': {...}, 'liab': {...}, 'equity': {...}}
    兼容真实报表：整词表头、多行换行、行次列、名称拆行、金额与名称错行。
    修复：所有者权益段的标题在pdf里常被拆成两行且用半角括号
    （如 "所有者权益(或股东权" + "益)："），旧逻辑识别不到标题，
    导致权益科目全部落进负债段。现按行距把同一单元格的换行拼回整名。"""
    centers, start, rowno_xs = _bs_header(rows)
    rowno_rights = _measure_rownum_rights(rows, start, rowno_xs)
    data = {'asset': {}, 'liab': {}, 'equity': {}}
    section = 'liab'
    # 左右两侧各自的状态：待拼名、待拼名所在行y、无名数值暂存、上一科目、上一科目行y
    st = {'L': {'pend': None, 'py': None, 'buf': None, 'last': None, 'ly': None},
          'R': {'pend': None, 'py': None, 'buf': None, 'last': None, 'ly': None}}

    def put(sec, side, name, raw, y):
        nk = norm_key(name)
        if not nk:
            return
        data[sec][nk] = {'val': parse_num(raw), 'raw': raw, 'name': name}
        st[side]['last'] = nk
        st[side]['ly'] = y

    def flush(sec, side):
        """把"有名称、值是下一行数值"的暂存科目写入"""
        s = st[side]
        if s['pend'] and s['buf'] is not None:
            put(sec, side, s['pend'], s['buf'], s['py'])
        s['pend'] = s['py'] = s['buf'] = None

    def append_to_last(sec, side, frag, y):
        """名称换行时把尾巴补回上一个科目名（如 "以公允价值计量且其变动
        计入当期损益的金" + "融资产"）"""
        s = st[side]
        nk = s['last']
        old = data[sec].get(nk) if nk else None
        if not old or not frag:
            return False
        newname = old['name'] + frag
        nnk = norm_key(newname)
        data[sec][nnk] = {'val': old['val'], 'raw': old['raw'], 'name': newname}
        del data[sec][nk]
        s['last'] = nnk
        s['ly'] = y
        return True

    for d in rows[start:]:
        y = d['y']
        name_a_parts, name_l_parts, vals = [], [], []
        for x, x1, t in d['tokens']:
            if NUM_TK.match(t.replace(' ', '')):
                if _is_rownum_tok(x, x1, rowno_xs):
                    continue  # 行次号，排除
                if x < 160:
                    # 无行次表头的报表（一般企业）：行号并入科目名
                    name_a_parts.append(t)
                    continue
                sp = _split_merged_rownum(x, x1, t, rowno_xs, rowno_rights)
                if sp is not None:
                    vals.append((sp[0], x1, sp[1]))  # 拆出行号后保留数值部分
                    continue
                vals.append((x, x1, t))
            elif x < 300:
                name_a_parts.append(t)
            else:
                name_l_parts.append(t)
        name_a = ''.join(name_a_parts).strip()
        name_l = ''.join(name_l_parts).strip()
        cols = _assign_by_center(vals, centers)
        av = ''.join(cols.get(centers[0], []))
        lv = ''.join(cols.get(centers[2], []))
        # ---- 资产侧 ----
        s = st['L']
        if name_a:
            if s['pend'] is not None and y - s['py'] <= NAME_LINE_GAP:
                name_a = s['pend'] + name_a      # 名称换行：拼回整名
                s['pend'] = s['py'] = None
            elif s['pend'] is not None:
                flush('asset', 'L')
            if not av and s['buf'] is not None:
                av = s['buf']                     # 数值出现在名称前一行
                s['buf'] = None
            if av:
                put('asset', 'L', name_a, av, y)
            elif s['last'] and y - s['ly'] <= NAME_LINE_GAP \
                    and not _not_item_name(norm_key(name_a)):
                append_to_last('asset', 'L', name_a, y)   # 名称尾巴（换行）
            else:
                s['pend'] = name_a
                s['py'] = y
        elif av:
            if s['pend'] is not None and y - s['py'] <= NAME_LINE_GAP:
                put('asset', 'L', s['pend'], av, y)
                s['pend'] = s['py'] = None
            else:
                if s['pend'] is not None:
                    flush('asset', 'L')
                s['buf'] = av
        # ---- 负债/权益侧 ----
        s = st['R']
        if name_l:
            if s['pend'] is not None:
                merge = (y - s['py'] <= NAME_LINE_GAP) or \
                        (bool(lv) and norm_key(name_l) in ('总计', '合计'))
                if merge:
                    name_l = s['pend'] + name_l
                    s['pend'] = s['py'] = None
                else:
                    flush(section, 'R')
            if not lv and s['buf'] is not None:
                lv = s['buf']                     # 数值出现在名称前一行
                s['buf'] = None
            if _is_section_header(norm_key(name_l)):
                section = 'equity'                # 进入"所有者权益"段
                s['last'] = s['ly'] = None
                continue
            if lv:
                put(section, 'R', name_l, lv, y)
            elif s['last'] and y - s['ly'] <= NAME_LINE_GAP \
                    and not _not_item_name(norm_key(name_l)):
                if not append_to_last(section, 'R', name_l, y):
                    s['pend'] = name_l
                    s['py'] = y
            else:
                s['pend'] = name_l
                s['py'] = y
        elif lv:
            if s['pend'] is not None and y - s['py'] <= NAME_LINE_GAP:
                put(section, 'R', s['pend'], lv, y)
                s['pend'] = s['py'] = None
            else:
                if s['pend'] is not None:
                    flush(section, 'R')
                s['buf'] = lv
    return data


def _parse_pl_pdf_page(rows):
    """解析利润表页 → {'pl': {normkey: {val(cur列), cum(累计列), name}}}
    列含义按表头文字动态识别：含"累计"字样 → cum，其余 → cur"""
    cols_hdr, start, rowno_xs = _pl_header(rows)
    rowno_rights = _measure_rownum_rights(rows, start, rowno_xs)
    centers = [c[0] for c in cols_hdr]
    kind_of = {c[0]: c[1] for c in cols_hdr}
    cur_c = next((c for c, k in cols_hdr if k == 'cur'), centers[0] if centers else 361)
    cum_c = next((c for c, k in cols_hdr if k == 'cum'), None)
    name_limit = min(centers) - 150 if centers else 211
    data = {'pl': {}}
    pending = None
    last_key = None
    for d in rows[start:]:
        name_items, vals = [], []
        for x, x1, t in d['tokens']:
            if NUM_TK.match(t.replace(' ', '')):
                if _is_rownum_tok(x, x1, rowno_xs):
                    continue  # 行次号，排除
                if x < name_limit:
                    # 无行次表头的报表（一般企业）：行号并入科目名
                    name_items.append((x, t))
                    continue
                sp = _split_merged_rownum(x, x1, t, rowno_xs, rowno_rights)
                if sp is not None:
                    vals.append((sp[0], x1, sp[1]))  # 拆出行号后保留数值部分
                    continue
                vals.append((x, x1, t))
            else:
                name_items.append((x, t))
        name = ''.join(t for _, t in sorted(name_items, key=lambda p: p[0])).strip()
        cols = _assign_by_center(vals, centers)
        cur = ''.join(cols.get(cur_c, []))
        cum = ''.join(cols.get(cum_c, [])) if cum_c else ''
        if name and (cur or cum):
            full = name
            if pending and '……' not in pending and norm_key(pending) != norm_key(full):
                data['pl'].setdefault(norm_key(pending), {
                    'val': None, 'raw': '', 'name': pending})
            pending = None
            data['pl'][norm_key(full)] = {
                'val': parse_num(cur), 'raw': cur, 'name': full,
                'cum': parse_num(cum) if cum else None}
            last_key = norm_key(full)
        elif name:
            if last_key and data['pl'][last_key]['name'].endswith('（损失'):
                old = data['pl'][last_key]
                newname = old['name'] + name
                nk = norm_key(newname)
                data['pl'][nk] = {'val': old['val'], 'raw': old['raw'],
                                  'name': newname, 'cum': old.get('cum')}
                del data['pl'][last_key]
                last_key = nk
                continue
            pending = (pending or '') + name
        elif (cur or cum) and pending:
            data['pl'][norm_key(pending)] = {
                'val': parse_num(cur), 'raw': cur, 'name': pending,
                'cum': parse_num(cum) if cum else None}
            last_key = norm_key(pending)
            pending = None
    return data


def read_pdf(path, kind):
    """读取财务报表pdf；kind='bs'取资产负债表页，kind='pl'取利润表页"""
    import pymupdf
    doc = pymupdf.open(path)
    try:
        for pi in range(len(doc)):
            rows = _rows_from_page(doc[pi])
            if not rows:
                continue
            txt = ''.join(t for d in rows for _, _, t in d['tokens'])
            if kind == 'bs' and '资产负债表' in txt:
                return _parse_bs_pdf_page(rows)
            if kind == 'pl' and '利润表' in txt:
                return _parse_pl_pdf_page(rows)
        return None
    finally:
        doc.close()


# ---------------------------------------------------------------- 匹配与填报

def lookup_item(src, side, tpl_item, aliases):
    """在源数据中查找模板科目：先别名，再精确，再去掉"加：/减："前缀回退"""
    nk = norm_key(tpl_item)
    cands = aliases.get(nk, []) + [nk]
    stripped = re.sub(r'^(加|减)[：:]', '', nk)
    if stripped != nk:
        cands.append(stripped)
    for cand in cands:
        if side in src and cand in src[side]:
            return src[side][cand]
    return None


def generate(current_xlsx, annual_pdf, quarter_pdf=None, small_prev_pdf=None,
             is_q1=None, progress_cb=None):
    """执行填报。
    quarter_pdf 上传 → 一般企业准则；small_prev_pdf 上传 → 小企业准则。
    is_q1：小企业准则下是否第一季度（None=自动按文件名年份判断，判断不出抛异常）"""
    aliases = load_aliases()
    if quarter_pdf and small_prev_pdf:
        raise ValueError('请只选择一种申报方式：一般企业（上传上年X季度财报）或'
                         '小企业（上传小企业上季度财报pdf），不要同时上传')
    if not quarter_pdf and not small_prev_pdf:
        raise ValueError('请上传"上年X季度财务报表"（一般企业）或'
                         '"小企业上季度财务报表pdf"（小企业准则）')
    kind = 'general' if quarter_pdf else 'small'
    spec = read_template_spec(kind)

    if progress_cb:
        progress_cb('读取当期excel...')
    n04 = read_n04(current_xlsx)
    if progress_cb:
        progress_cb('读取上年年报/第四季度pdf...')
    annual = read_pdf(annual_pdf, 'bs')
    if annual is None:
        raise ValueError('未能从UI2 pdf中识别资产负债表页')
    quarter = None
    small_prev = None
    if kind == 'general':
        if progress_cb:
            progress_cb('读取上年X季度pdf...')
        quarter = read_pdf(quarter_pdf, 'pl')
        if quarter is None:
            raise ValueError('未能从上年X季度pdf中识别利润表页')
    else:
        if progress_cb:
            progress_cb('读取小企业上季度财报pdf...')
        prev_pl = read_pdf(small_prev_pdf, 'pl')
        if prev_pl is None:
            raise ValueError('未能从小企业上季度pdf中识别利润表页')
        # 小企业利润表取"本年累计金额"列（解析结果中的cum列）
        small_prev = {'pl': {}}
        for k, v in prev_pl['pl'].items():
            small_prev['pl'][k] = {'val': v.get('cum'), 'raw': v.get('cum'),
                                   'name': v['name']}
        if is_q1 is None:
            y1 = extract_year(os.path.basename(current_xlsx))
            y4 = extract_year(os.path.basename(small_prev_pdf))
            if y1 is None or y4 is None:
                raise ValueError('无法从文件名判断年份，请确认UI1与UI4文件名包含20xx年份，'
                                 '或在界面上选择季度')
            is_q1 = (y1 != y4)

    # 生成填报值
    fillmap = {}
    logs = []
    for sp in spec:
        item, side, ftype = sp['item'], sp['side'], sp['fill_type']
        status = '未找到源科目'
        val = None
        src_name = ''
        src_item = ''
        note = ''
        if ftype == 'current':
            src, src_name = n04, '当期excel'
        elif ftype == 'annual':
            src, src_name = annual, 'UI2 pdf'
        elif ftype == 'quarter_prev':
            src, src_name = quarter, '上年X季度pdf'
        else:
            src, src_name = None, None  # small_calc 单独处理
        if ftype == 'small_calc':
            cur_hit = lookup_item(n04, side, item, aliases)
            prev_hit = lookup_item(small_prev, side, item, aliases)
            if cur_hit is not None and cur_hit['val'] is not None:
                try:
                    cur = float(cur_hit['val'])
                except (TypeError, ValueError):
                    cur = None
                if cur is None:
                    status = '源数据非数值'
                    note = '当期excel该科目数值无法识别：%s' % cur_hit['val']
                    src_name = '当期excel'
                    src_item = cur_hit['name']
                elif is_q1:
                    val = cur
                    status = '已填入(第一季度)'
                    note = '本期金额=本年累计金额'
                    src_name = '当期excel'
                    src_item = cur_hit['name']
                elif prev_hit is not None and prev_hit['val'] is not None:
                    try:
                        prev = float(prev_hit['val'])
                    except (TypeError, ValueError):
                        prev = None
                    if prev is None:
                        status = '源数据非数值'
                        note = '上期pdf该科目数值无法识别：%s' % prev_hit['val']
                        src_name = '上期pdf'
                        src_item = prev_hit['name']
                    else:
                        val = round(cur - prev, 2)
                        status = '已填入(计算)'
                        note = '本期金额=%s-%s(当期累计-上期累计)' % (cur, prev)
                        src_name = '当期excel-上期pdf'
                        src_item = '%s / %s' % (cur_hit['name'], prev_hit['name'])
                else:
                    status = '上期数据缺失'
                    note = '当期本年累计=%s，但上期pdf中无该科目本年累计数' % cur
                    src_name = '当期excel'
                    src_item = cur_hit['name']
            elif cur_hit is not None:
                status = '源科目为空'
                src_name = '当期excel'
                src_item = cur_hit['name']
            else:
                status = '未找到源科目'
        else:
            hit = lookup_item(src, side, item, aliases)
            if hit is not None:
                src_item = hit['name']
                if hit['val'] is None:
                    status = '源科目为空'
                else:
                    try:
                        val = float(hit['val'])
                        status = '已填入'
                    except (TypeError, ValueError):
                        status = '源数据非数值'
        key = (sp['sheet'], sp['row0'], sp['col0'])
        fillmap[key] = val
        logs.append({'sheet': sp['sheet'], 'row': sp['row0'] + 1,
                     'col': sp['col0'] + 1, 'colname': sp['colname'],
                     'item': item, 'value': val, 'src': src_name,
                     'src_item': src_item, 'status': status, 'note': note})

    # ---- 小企业利润表：信用减值损失重分类 + 勾稽计算 + 勾稽校验 ----
    pl_checks = []
    if kind == 'small':
        try:
            eqs = small_pl_equations(get_template_path('small'))
        except Exception:
            eqs = None
        if eqs is None:
            pl_checks.append('小企业模板利润表结构无法识别，未做勾稽计算与校验')
        sheet_pl = next((sp['sheet'] for sp in spec if sp['side'] == 'pl'), None)
        # 当期excel是否存在信用减值损失（小企业准则无此科目，需重分类至营业外支出）
        imp = None
        if eqs is not None:
            for k, v in n04['pl'].items():
                if '信用减值损失' in k and v.get('val') is not None:
                    try:
                        imp = {'val': float(v['val']), 'name': v['name']}
                    except (TypeError, ValueError):
                        imp = None
                    if imp is not None:
                        break
        # 模板若自带信用减值损失行则按原映射填报，无需重分类
        tpl_has_imp_row = eqs is not None and any(
            '信用减值损失' in norm_key(sp['item'])
            for sp in spec if sp['side'] == 'pl')
        rec_by_key = {(lg['sheet'], lg['row'] - 1, lg['col'] - 1): lg for lg in logs}
        # 行号 → {col0: 值, ('st',col0): 状态, item: 科目}
        pl_recs = {}
        for sp in spec:
            if sp['side'] != 'pl':
                continue
            key = (sp['sheet'], sp['row0'], sp['col0'])
            d = pl_recs.setdefault(sp['row0'], {'item': sp['item']})
            d[sp['col0']] = fillmap.get(key)
            d[('st', sp['col0'])] = rec_by_key[key]['status']

        if eqs is not None and imp is not None and not tpl_has_imp_row:
            imp_abs = abs(imp['val'])
            # 当期excel自身的营业外支出（如有，与重分类金额合并）
            excel_oe = excel_oe_name = None
            for k, v in n04['pl'].items():
                if '营业外支出' in k and '其中' not in k and v.get('val') is not None:
                    try:
                        excel_oe = float(v['val'])
                    except (TypeError, ValueError):
                        excel_oe = None
                    if excel_oe is not None:
                        excel_oe_name = v['name']
                        break
            oe_cum = round(imp_abs + (excel_oe or 0.0), 2)
            plus_txt = ' + 营业外支出%s(%s)' % (excel_oe, excel_oe_name) \
                if excel_oe is not None else ''
            for sp in spec:
                if sp['side'] != 'pl' or \
                        '营业外支出' not in norm_key(sp['item']):
                    continue
                key = (sp['sheet'], sp['row0'], sp['col0'])
                lg = rec_by_key[key]
                if sp['fill_type'] == 'current':      # 本年累计金额列
                    fillmap[key] = oe_cum
                    lg.update(value=oe_cum, status='已填入(重分类)',
                              src='当期excel(信用减值损失重分类)',
                              src_item=imp['name'] + ((' + ' + excel_oe_name)
                                                      if excel_oe is not None else ''),
                              note='小企业准则：信用减值损失计入营业外支出'
                                   '：|%s|%s = %s' % (imp['val'], plus_txt, oe_cum))
                    pl_recs[sp['row0']][sp['col0']] = oe_cum
                    pl_recs[sp['row0']][('st', sp['col0'])] = lg['status']
                else:                                  # 本期金额列(small_calc)
                    if is_q1:
                        fillmap[key] = oe_cum
                        lg.update(value=oe_cum, status='已填入(重分类)',
                                  src='当期excel(信用减值损失重分类)',
                                  src_item=imp['name'],
                                  note='第一季度：本期金额=本年累计金额(%s)' % oe_cum)
                        pl_recs[sp['row0']][sp['col0']] = oe_cum
                    else:
                        prev_hit = lookup_item(small_prev, 'pl', sp['item'], aliases)
                        prev_v = None
                        if prev_hit is not None and prev_hit.get('val') is not None:
                            try:
                                prev_v = float(prev_hit['val'])
                            except (TypeError, ValueError):
                                prev_v = None
                        if prev_v is None:
                            lg.update(status='上期数据缺失',
                                      src='当期excel(信用减值损失重分类)',
                                      src_item=imp['name'],
                                      note='重分类后本年累计营业外支出=%s，'
                                           '但上期pdf无营业外支出本年累计数' % oe_cum)
                            pl_recs[sp['row0']][('st', sp['col0'])] = '上期数据缺失'
                        else:
                            q = round(oe_cum - prev_v, 2)
                            fillmap[key] = q
                            lg.update(value=q, status='已填入(重分类计算)',
                                      src='当期excel-上期pdf(信用减值损失重分类)',
                                      src_item='%s / %s' % (imp['name'], prev_hit['name']),
                                      note='本期营业外支出=|%s|%s-%s=%s'
                                           '（本季度新增信用减值损失）'
                                           % (imp['val'], plus_txt, prev_v, q))
                            pl_recs[sp['row0']][sp['col0']] = q
                            pl_recs[sp['row0']][('st', sp['col0'])] = lg['status']
                            if q < 0:
                                pl_checks.append(
                                    '利润表[本期金额] 营业外支出=%s 为负数'
                                    '（信用减值损失转回？），请人工确认' % q)
                    pl_recs[sp['row0']][('st', sp['col0'])] = lg['status']

            # 营业利润/利润总额/净利润：按模板上行科目勾稽计算（不做两表直接相减）
            def seg_sum(seg, col):
                s = 0.0
                for r0, sg, _it in seg:
                    v = pl_recs.get(r0, {}).get(col)
                    if v is not None:
                        s += sg * float(v)
                return s

            def seg_note(seg, col):
                parts = []
                for r0, sg, it in seg:
                    nm = re.sub(r'^(减|加)[：:]', '', it).strip()
                    v = pl_recs.get(r0, {}).get(col)
                    parts.append('%s%s%s' % ('+' if sg > 0 else '-', nm,
                                             ('(=%s)' % round(float(v), 2))
                                             if v is not None else '(空)'))
                return ''.join(parts)

            colname_of = {3: '本期金额', 4: '本年累计金额'}
            for col in (3, 4):
                rev_v = pl_recs.get(eqs['totals']['rev'], {}).get(col)
                op_v = round((rev_v or 0.0) + seg_sum(eqs['segA'], col), 2)
                tp_v = round(op_v + seg_sum(eqs['segB'], col), 2)
                nt_v = round(tp_v + seg_sum(eqs['segC'], col), 2)
                for r0, tv, txt in (
                        (eqs['totals']['op'], op_v,
                         '勾稽：营业收入(=%s)%s = %s（信用减值损失%s已重分类至'
                         '营业外支出，不影响营业利润）'
                         % ('' if rev_v is None else round(float(rev_v), 2),
                            seg_note(eqs['segA'], col), op_v, imp['val'])),
                        (eqs['totals']['total'], tp_v,
                         '勾稽：营业利润%s%s = %s' % (op_v,
                                                      seg_note(eqs['segB'], col), tp_v)),
                        (eqs['totals']['net'], nt_v,
                         '勾稽：利润总额%s%s = %s' % (tp_v,
                                                      seg_note(eqs['segC'], col), nt_v))):
                    key = (sheet_pl, r0, col)
                    fillmap[key] = tv
                    lg = rec_by_key.get(key)
                    if lg is not None:
                        lg.update(value=tv, status='已填入(勾稽计算)',
                                  src='按模板勾稽计算',
                                  src_item='上行收入成本费用科目',
                                  note=txt)
                    pl_recs[r0][col] = tv
                    pl_recs[r0][('st', col)] = '已填入(勾稽计算)'

            # 信用减值损失重分类后的整体勾稽：模板利润总额(累计) 应等于 excel利润总额
            xl_tp = None
            for k, v in n04['pl'].items():
                if k.startswith('三、利润总额') and v.get('val') is not None:
                    try:
                        xl_tp = float(v['val'])
                    except (TypeError, ValueError):
                        xl_tp = None
                    if xl_tp is not None:
                        break
            tpl_tp = pl_recs.get(eqs['totals']['total'], {}).get(4)
            if xl_tp is not None and tpl_tp is not None:
                diff = round(tpl_tp - xl_tp, 2)
                if abs(diff) > 0.005:
                    pl_checks.append(
                        '利润表[本年累计金额] 模板利润总额(勾稽)%s 与 当期excel利润总额%s '
                        '不一致，差%s，请检查信用减值损失重分类或excel营业外支出'
                        % (tpl_tp, xl_tp, diff))
                else:
                    pl_checks.append(
                        '利润表[本年累计金额] 模板利润总额(勾稽)%s 与 当期excel利润总额%s '
                        '一致，信用减值损失重分类后勾稽成立' % (tpl_tp, xl_tp))

        # ---- 勾稽校验（小企业利润表两列：营业利润/利润总额/净利润 vs 上行科目）----
        if eqs is not None:
            unknown_st = ('未找到源科目', '上期数据缺失', '源数据非数值')
            colname_of = {3: '本期金额', 4: '本年累计金额'}
            chain = [('营业利润', eqs['totals']['op'], eqs['segA'],
                      eqs['totals']['rev']),
                     ('利润总额', eqs['totals']['total'], eqs['segB'],
                      eqs['totals']['op']),
                     ('净利润', eqs['totals']['net'], eqs['segC'],
                      eqs['totals']['total'])]
            for col in (3, 4):
                for tname, trow, cseg, brow in chain:
                    tv = pl_recs.get(trow, {}).get(col)
                    if tv is None:
                        continue
                    comp = 0.0
                    missing = []
                    for r0, sg, it in cseg:
                        v = pl_recs.get(r0, {}).get(col)
                        if v is None:
                            if pl_recs.get(r0, {}).get(('st', col)) in unknown_st:
                                missing.append(it.strip())
                        else:
                            comp += sg * float(v)
                    if brow is not None:
                        bv = pl_recs.get(brow, {}).get(col)
                        if bv is None:
                            if pl_recs.get(brow, {}).get(('st', col)) in unknown_st:
                                missing.append('上级小计')
                        else:
                            comp += float(bv)
                    tag = '利润表[%s] %s' % (colname_of[col], tname)
                    if missing:
                        pl_checks.append('%s：勾稽校验跳过（未取得：%s）'
                                         % (tag, '、'.join(missing[:4])))
                        continue
                    comp = round(comp, 2)
                    diff = round(float(tv) - comp, 2)
                    if abs(diff) > 0.005:
                        pl_checks.append('%s：填入%s ≠ 上行勾稽%s，差%s，请检查填报！'
                                         % (tag, tv, comp, diff))
                    else:
                        pl_checks.append('%s：勾稽一致（%s）' % (tag, tv))

    # 写入模板副本：保留原调色板（修复颜色变暗问题），不填充任何单元格格式
    import xlrd
    from xlutils.copy import copy as xlcopy
    tpath = get_template_path(kind)
    rb = xlrd.open_workbook(tpath, formatting_info=True)
    wb = xlcopy(rb)
    # 3.0修复1：把模板的自定义调色板复制到输出文件
    for idx, rgb in rb.colour_map.items():
        try:
            wb.set_colour_RGB(idx, *rgb)
        except Exception:
            pass
    for (sheet, r, c), val in fillmap.items():
        if val is not None:
            wb.get_sheet(sheet).write(r, c, val)
    os.makedirs(OUT_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    prefix = '一般企业' if kind == 'general' else '小企业会计准则'
    outpath = os.path.join(OUT_DIR, '财务报表报送与信息采集（%s）_填报结果_%s.xls'
                           % (prefix, ts))
    wb.save(outpath)

    if kind == 'small':
        qdesc = '第一季度：本期金额=本年累计金额' if is_q1 \
            else '第2~4季度：本期金额=当期本年累计-上期本年累计(上期pdf本年累计金额列)'
    else:
        qdesc = '一般企业准则：本期金额=当期本期发生额，上期金额=上年同期本期金额'
    checks = []
    for sp in spec:
        if sp['side'] in ('asset', 'liab', 'equity') and sp['fill_type'] == 'annual':
            a_hit = lookup_item(annual, sp['side'], sp['item'], aliases)
            n_hit = lookup_item(n04, sp['side'], sp['item'], aliases)
            if a_hit and a_hit['val'] is not None and n_hit and n_hit.get('begin') is not None:
                try:
                    av = float(a_hit['val'])
                    nb = float(n_hit['begin'])
                    if abs(av - nb) > 0.005:
                        checks.append('差值: %s | UI2 pdf期末=%s | 当期excel期初=%s'
                                      % (sp['item'], av, nb))
                except (TypeError, ValueError):
                    pass
    logpath = write_log(logs, checks, qdesc, kind,
                        current_xlsx, annual_pdf, quarter_pdf, small_prev_pdf,
                        is_q1 if kind == 'small' else None, pl_checks)
    return outpath, logpath, logs, checks, kind, pl_checks


def write_log(logs, checks, qdesc, kind, cur, annual, quarter, small_prev, is_q1,
              pl_checks=None):
    """写校验日志xlsx；pl_checks：小企业利润表勾稽校验结果"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    os.makedirs(OUT_DIR, exist_ok=True)
    ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    path = os.path.join(OUT_DIR, '填报校验日志_%s.xlsx' % ts)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '填报明细'
    headers = ['模板sheet', '模板行', '模板列', '列含义', '模板科目', '填入值',
               '数据来源', '来源科目', '状态', '说明']
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='D9E1F2')
    red = PatternFill('solid', fgColor='FFC7CE')
    yellow = PatternFill('solid', fgColor='FFEB9C')
    green = PatternFill('solid', fgColor='C6EFCE')
    for i, lg in enumerate(logs, 2):
        vals = [lg['sheet'], lg['row'], lg['col'], lg['colname'], lg['item'],
                lg['value'], lg['src'], lg['src_item'], lg['status'], lg['note']]
        for j, v in enumerate(vals, 1):
            c = ws.cell(row=i, column=j, value=v)
            if j == 9:
                if v in ('未找到源科目', '上期数据缺失', '源数据非数值'):
                    c.fill = red
                elif v == '源科目为空':
                    c.fill = yellow
                elif v and v.startswith('已填入'):
                    c.fill = green
    widths = [11, 7, 7, 11, 32, 16, 14, 28, 12, 34]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws2 = wb.create_sheet('汇总')
    cnt = {}
    for lg in logs:
        cnt[lg['status']] = cnt.get(lg['status'], 0) + 1
    rows = [('生成时间', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
            ('适用准则', '一般企业' if kind == 'general' else '小企业会计准则'),
            ('填报逻辑', qdesc),
            ('当期excel', cur),
            ('UI2 上年年报/第四季度pdf', annual),
            ('UI3 上年X季度pdf', quarter or '（未上传，小企业准则）'),
            ('UI4 小企业上季度财报pdf', small_prev or '（未上传，一般企业准则）'),
            ('是否按第一季度', {None: '不适用', True: '是', False: '否'}[is_q1]),
            ('', '')] + [('状态：' + k, v) for k, v in sorted(cnt.items())] + \
        [('', ''), ('交叉校验：当期excel期初余额 与 UI2 pdf期末余额 对比', '')]
    if checks:
        rows += [('不一致科目', c) for c in checks]
    else:
        rows.append(('不一致科目', '无（全部一致）'))
    if pl_checks:
        rows.append(('', ''))
        rows.append(('勾稽校验：小企业利润表', ''))
        rows += [('勾稽校验', c) for c in pl_checks]
    for i, (k, v) in enumerate(rows, 1):
        ws2.cell(row=i, column=1, value=k).font = Font(bold=(i <= 8 or v == ''))
        ws2.cell(row=i, column=2, value=v)
    ws2.column_dimensions['A'].width = 26
    ws2.column_dimensions['B'].width = 80
    wb.save(path)
    return path


# ---------------------------------------------------------------- 命令行

def run_cli(args):
    sys.stdout.reconfigure(encoding='utf-8')
    current, annual = args[0], args[1]
    ui3 = ui4 = None
    is_q1 = None
    i = 2
    while i < len(args):
        if args[i] == '--ui3':
            ui3 = args[i + 1]
        elif args[i] == '--ui4':
            ui4 = args[i + 1]
        elif args[i] == '--q1':
            is_q1 = args[i + 1].lower() in ('yes', 'true', '1')
        i += 2
    out, log, records, checks, kind, pl_checks = generate(
        current, annual, quarter_pdf=ui3, small_prev_pdf=ui4, is_q1=is_q1)
    print('适用准则:', '一般企业' if kind == 'general' else '小企业会计准则')
    print('填报结果：%s' % out)
    print('校验日志：%s' % log)
    cnt = {}
    for r in records:
        cnt[r['status']] = cnt.get(r['status'], 0) + 1
    print('状态统计：', cnt)
    print('交叉校验不一致科目数：', len(checks))
    for c in checks:
        print('  !', c)
    if pl_checks:
        print('勾稽校验（小企业利润表）：')
        for c in pl_checks:
            print('  -', c)


# ---------------------------------------------------------------- GUI

def main_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    root = tk.Tk()
    root.title('财务报表智能填报工具 3.0')
    root.geometry('1040x700')
    pad = {'padx': 8, 'pady': 3}

    frm = ttk.LabelFrame(root, text='上传文件（UI1、UI2必选；UI3与UI4二选一）')
    frm.pack(fill='x', **pad)
    labels = ['此处上传NOX-X-某主体202x0x（当期）excel文件',
              '此处上传上年年度财务报表或第四季度财务报表pdf文件',
              '此处上传上年X季度财务报表（一般企业申报用）',
              '此处上传小企业上季度财务报表pdf文件（小企业申报用）']
    kinds = [('Excel文件', '*.xlsx'), ('PDF文件', '*.pdf'),
             ('PDF文件', '*.pdf'), ('PDF文件', '*.pdf')]
    vars_ = []
    for i, (txt, (ft, pat)) in enumerate(zip(labels, kinds)):
        ttk.Label(frm, text=txt).grid(row=i, column=0, sticky='w', **pad)
        var = tk.StringVar()
        ttk.Entry(frm, textvariable=var, width=76).grid(row=i, column=1, **pad)

        def browse(v=var, ft=ft, pat=pat):
            p = filedialog.askopenfilename(
                title='选择文件', filetypes=[(ft, pat), ('所有文件', '*.*')])
            if p:
                v.set(p)
        ttk.Button(frm, text='选择文件…', command=browse).grid(row=i, column=2, **pad)
        vars_.append(var)

    frm_t = ttk.Frame(root)
    frm_t.pack(fill='x', **pad)
    ttk.Button(frm_t, text='替换模板…', command=lambda: replace_template(root)).pack(
        side='left')
    ttk.Button(frm_t, text='打开模板文件夹', command=lambda: open_dir(TPL_DIR)).pack(
        side='left', padx=8)
    ttk.Button(frm_t, text='打开填报结果文件夹', command=lambda: open_dir(OUT_DIR)).pack(
        side='left')
    ttk.Button(frm_t, text='生成填报结果', command=lambda: do_generate(root, vars_)).pack(
        side='right')
    lbl_status = ttk.Label(frm_t, text='')
    lbl_status.pack(side='right', padx=12)

    frm_rs = ttk.LabelFrame(root, text='填报预览（模板科目 → 填入值 / 状态）')
    frm_rs.pack(fill='both', expand=True, **pad)
    tree = ttk.Treeview(frm_rs, columns=('sheet', 'item', 'colname', 'val',
                                         'src', 'status', 'note'),
                        show='headings')
    for c, txt, w in [('sheet', '表', 90), ('item', '模板科目', 260),
                      ('colname', '列', 90), ('val', '填入值', 120),
                      ('src', '来源', 130), ('status', '状态', 110),
                      ('note', '说明', 200)]:
        tree.heading(c, text=txt)
        tree.column(c, width=w)
    tree.pack(side='left', fill='both', expand=True)
    vsb = ttk.Scrollbar(frm_rs, orient='vertical', command=tree.yview)
    vsb.pack(side='right', fill='y')
    tree.configure(yscrollcommand=vsb.set)
    tree.tag_configure('ok', foreground='#1e7b1e')
    tree.tag_configure('bad', foreground='#c00000')
    tree.tag_configure('warn', foreground='#b8860b')

    def do_generate(r, vs):
        paths = [v.get().strip() for v in vs]
        if not paths[0] or not paths[1]:
            messagebox.showwarning('提示', '请先选择 UI1（当期excel）和 UI2（上年年报/第四季度pdf）')
            return
        has3 = bool(paths[2])
        has4 = bool(paths[3])
        if has3 == has4:
            messagebox.showwarning('提示',
                                   'UI3与UI4必须二选一：\n一般企业申报选 UI3（上年X季度财报pdf）；\n'
                                   '小企业申报选 UI4（小企业上季度财报pdf）。')
            return
        for p in paths:
            if p and not os.path.exists(p):
                messagebox.showerror('错误', '文件不存在：%s' % p)
                return
        is_q1 = None
        if has4:
            y1 = extract_year(os.path.basename(paths[0]))
            y4 = extract_year(os.path.basename(paths[3]))
            if y1 is None or y4 is None:
                ans = messagebox.askyesno('确认季度',
                                          '无法从文件名自动判断年份。\n\n'
                                          '是否按第一季度申报？\n'
                                          '（是：本期金额=本年累计金额；\n'
                                          '  否：本期金额=当期累计-上期累计）')
                is_q1 = ans
            else:
                is_q1 = (y1 != y4)
        tree.delete(*tree.get_children())
        lbl_status.config(text='正在填报…')
        r.config(cursor='wait')
        r.update()
        try:
            out, log, records, checks, kind, pl_checks = generate(
                paths[0], paths[1], quarter_pdf=paths[2] or None,
                small_prev_pdf=paths[3] or None, is_q1=is_q1)
        except Exception:
            traceback.print_exc()
            r.config(cursor='')
            lbl_status.config(text='填报失败')
            messagebox.showerror('错误', '填报失败：\n%s' % traceback.format_exc(limit=2))
            return
        r.config(cursor='')
        tagmap = {'已填入': 'ok', '已填入(计算)': 'ok', '已填入(第一季度)': 'ok',
                  '已填入(勾稽计算)': 'ok', '已填入(重分类)': 'ok',
                  '已填入(重分类计算)': 'ok',
                  '未找到源科目': 'bad', '上期数据缺失': 'bad', '源数据非数值': 'bad',
                  '源科目为空': 'warn'}
        for rec in records:
            tree.insert('', 'end', values=(
                rec['sheet'], rec['item'], rec['colname'],
                '' if rec['value'] is None else rec['value'],
                rec['src'] + (' / ' + rec['src_item'][:14] if rec['src_item'] else ''),
                rec['status'], rec['note']), tags=(tagmap.get(rec['status'], ''),))
        lbl_status.config(text='完成：%s' % ('一般企业' if kind == 'general' else '小企业会计准则'))
        msg = '填报完成（%s）\n\n结果：%s\n日志：%s' % (
            '一般企业' if kind == 'general' else '小企业会计准则', out, log)
        if checks:
            msg += '\n\n⚠ 交叉校验发现 %d 个科目与期初不一致（详见日志）：\n' % len(checks)
            msg += '\n'.join(checks[:8])
        if pl_checks:
            bad = [c for c in pl_checks if ('≠' in c or '不一致' in c or '为负数' in c)]
            if bad:
                msg += '\n\n⚠ 勾稽校验发现 %d 项异常（详见日志）：\n' % len(bad)
                msg += '\n'.join(bad[:8])
            else:
                msg += '\n\n勾稽校验（利润表）：%d 项全部通过' % len(pl_checks)
        messagebox.showinfo('完成', msg)

    def replace_template(r):
        p = filedialog.askopenfilename(title='选择模板xls文件（文件名含"小企业"=小企业模板，否则为一般企业模板）',
                                       filetypes=[('Excel 97-2003模板', '*.xls')])
        if not p:
            return
        os.makedirs(TPL_DIR, exist_ok=True)
        is_small = '小企业' in os.path.basename(p)
        for f in os.listdir(TPL_DIR):
            f_small = '小企业' in f
            if f.lower().endswith('.xls') and f_small == is_small:
                os.remove(os.path.join(TPL_DIR, f))
        shutil.copy2(p, os.path.join(TPL_DIR, os.path.basename(p)))
        messagebox.showinfo('完成', '模板已替换：%s' % os.path.basename(p))

    def open_dir(d):
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    ex = [f for f in os.listdir(APP_DIR) if f.lower().endswith('.xlsx')
          and not f.startswith('~$') and re.search(r'20\d{2}(0[1-9]|1[0-2])', f)]
    if ex:
        vars_[0].set(os.path.join(APP_DIR, sorted(ex)[0]))
    pdfs = [f for f in os.listdir(APP_DIR) if f.lower().endswith('.pdf')]
    for f in sorted(pdfs):
        if '年度' in f or '第4季度' in f or '第四季度' in f:
            if not vars_[1].get():
                vars_[1].set(os.path.join(APP_DIR, f))
        elif '季度' in f:
            if not vars_[2].get():
                vars_[2].set(os.path.join(APP_DIR, f))
    root.mainloop()


if __name__ == '__main__':
    if '--cli' in sys.argv:
        args = sys.argv[sys.argv.index('--cli') + 1:]
        if len(args) < 2:
            print('用法: 智能整合填报3.0.py --cli <UI1当期excel> <UI2年报/第四季度pdf>'
                  ' [--ui3 上年X季度pdf | --ui4 小企业上季度pdf] [--q1 yes|no]')
            sys.exit(1)
        run_cli(args)
    else:
        main_gui()
