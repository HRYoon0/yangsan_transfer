# 학교관리자취합용 엑셀(xlsm)의 명부 9종 시트에서 서식(열 너비, 머리글, 글꼴, 테두리, 인쇄 설정)만 뽑아
# manager.html 의 ROSTER_SPEC 에 넣는다. 명부에 적힌 자료(개인정보)는 읽지 않는다.
#   사용법: python tools/roster_spec.py "<학교관리자취합용 xlsm 경로>"
import sys, os, json, re, openpyxl
from openpyxl.utils import get_column_letter as L
sys.stdout.reconfigure(encoding='utf-8')
# 시트 이름: (제목 행, 첫 자료 행)
ROSTERS = {'서열명부': (5, 9), '임용제청서': (5, 9), '승진후보자명부': (5, 10), '장기근속자명부': (5, 9), '타시도(군)전보내신자': (5, 9),
           '부부교직원': (5, 10), '유예신청자명부': (4, 9), '전보우선자명부': (5, 9), '과원전보내신자명부': (5, 9)}
def style(c):
    f, b, a = c.font, c.border, c.alignment
    return {'f': [f.name, f.sz, 1 if f.b else 0],
            'b': ''.join({'thin': 't', 'medium': 'm', 'hair': 'h', 'double': 'd'}.get(getattr(b, s).style, '-') for s in ('left', 'right', 'top', 'bottom')),
            'a': [a.horizontal or '', a.vertical or '', 1 if a.wrap_text else 0, 1 if a.shrink_to_fit else 0],
            'n': c.number_format}
def widths(ws, count):
    # 여러 열을 한 묶음으로 지정한 너비(min~max)를 열마다 풀어서 읽는다
    out = [None] * count
    for d in ws.column_dimensions.values():
        if d.width is None: continue
        for i in range(d.min or 0, (d.max or 0) + 1):
            if 1 <= i <= count and (out[i - 1] is None or d.customWidth): out[i - 1] = round(d.width, 2)
    default = ws.sheet_format.defaultColWidth or 8.43
    return [w if w is not None else round(default, 2) for w in out]
def extract(path):
    wb = openpyxl.load_workbook(path)
    styles, out = [], {}
    def sid(c):
        s = style(c)
        if s not in styles: styles.append(s)
        return styles.index(s)
    for name, (title_row, first) in ROSTERS.items():
        ws = wb[name]
        last = max(c.column for row in ws.iter_rows(min_row=title_row, max_row=first - 1) for c in row if c.value is not None)
        title = ws.cell(title_row, 1).value
        m = re.match(r'=\s*"(.*?)\("\s*&\s*메뉴!E2\s*&\s*"\)"', title)
        ps, pm = ws.page_setup, ws.page_margins
        head = []
        for r in range(title_row + 2, first):
            for c in range(1, last + 1):
                cell = ws.cell(r, c)
                head.append([r, c, cell.value if isinstance(cell.value, str) and not cell.value.startswith('=') else None, sid(cell)])
        out[name] = {
            'titleRow': title_row, 'first': first, 'cols': last, 'title': m.group(1),
            'widths': widths(ws, last), 'hiddenCols': sorted({i for d in ws.column_dimensions.values() if d.hidden for i in range(d.min, d.max + 1) if i <= last}), 'hiddenRows': [r for r in range(title_row, first + 3) if ws.row_dimensions[r].hidden], 'firstHeight': ws.row_dimensions[first].height,
            'heights': {r: ws.row_dimensions[r].height for r in range(1, first)}, 'dataHeight': ws.row_dimensions[first + 1].height,
            'merges': sorted(str(x) for x in ws.merged_cells.ranges if x.min_row >= title_row and x.max_row < first),
            'titleStyle': sid(ws.cell(title_row, 1)), 'principalStyle': sid(ws.cell(title_row + 1, 1)),
            'head': head, 'data': [sid(ws.cell(first + 1, c)) for c in range(1, last + 1)],
            'page': {'orientation': ps.orientation, 'paper': ps.paperSize, 'scale': ps.scale, 'hc': bool(ws.print_options.horizontalCentered),
                     'margins': [round(x, 4) for x in (pm.left, pm.right, pm.top, pm.bottom, pm.header, pm.footer)], 'titles': (ws.print_title_rows or '').replace('$', '')},
        }
    ws = wb['자료']
    data = {'widths': widths(ws, 33), 'merges': sorted(str(x) for x in ws.merged_cells.ranges),
            'row1': {c.coordinate: c.value for c in ws[1] if c.value is not None and c.column > 1}, 'row2': [c.value for c in ws[2][:33]],
            'fills': {c.coordinate: (c.fill.fgColor.rgb if isinstance(c.fill.fgColor.rgb, str) else 'theme%s' % c.fill.fgColor.theme) for c in ws[1][:33] if c.fill.fill_type}, 'h1': ws.row_dimensions[1].height}
    return {'styles': styles, 'rosters': out, 'data': data}
if __name__ == '__main__':
    spec = json.dumps(extract(sys.argv[1]), ensure_ascii=False, separators=(',', ':'))
    assert '</script' not in spec
    target = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'manager.html')
    lines = open(target, encoding='utf-8').read().split(chr(10))
    found = [i for i, line in enumerate(lines) if line.strip().startswith('const ROSTER_SPEC = ')]
    assert len(found) == 1, 'manager.html 에서 ROSTER_SPEC 줄을 찾지 못함'
    lines[found[0]] = '        const ROSTER_SPEC = ' + spec + ';'
    open(target, 'w', encoding='utf-8', newline=chr(10)).write(chr(10).join(lines))
    print('manager.html 의 명부 서식을 바꿨습니다.', len(spec), '글자')
