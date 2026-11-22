# 학교관리자취합용 엑셀(xlsm)의 명부 9종 시트에서 서식만 뽑아 manager.html 의 ROSTER_SPEC 에 넣는다.
# 명부에 적힌 자료(개인정보)는 읽지 않는다.
#   사용법: python tools/roster_spec.py "<학교관리자취합용 xlsm 경로>"
#
#   - 열 너비·행 높이·병합·인쇄 설정: 파일에서 읽음 (openpyxl)
#   - 칸마다의 글꼴·테두리·맞춤·표시 형식: 실제 엑셀에서 읽음 (pywin32)
#     (openpyxl 은 병합된 칸의 테두리를 스스로 채워서 읽기 때문에 원본과 달라진다. 그래서 엑셀이 보는 값을 그대로 쓴다.)
#   - 엑셀은 매크로를 끈 채 읽기 전용으로 열고, 화면에 띄우지 않으며, 끝나면 닫는다.
import sys, os, json, re, openpyxl
import win32com.client
sys.stdout.reconfigure(encoding='utf-8')

# 시트 이름: (제목 행, 첫 자료 행)
ROSTERS = {'서열명부': (5, 9), '임용제청서': (5, 9), '승진후보자명부': (5, 10), '장기근속자명부': (5, 9), '타시도(군)전보내신자': (5, 9),
           '부부교직원': (5, 10), '유예신청자명부': (4, 9), '전보우선자명부': (5, 9), '과원전보내신자명부': (5, 9)}
DATA_ROWS = 70      # 자료 줄은 이만큼까지 줄마다 서식을 읽는다 (원본은 줄에 따라 글꼴·높이가 조금씩 다르다)
H_ALIGN = {-4108: 'center', -4152: 'right', -4131: 'left', -4130: 'justify', 7: 'centerContinuous', 5: 'fill', -4117: 'distributed'}
V_ALIGN = {-4108: 'center', -4107: 'bottom', -4160: 'top', -4130: 'justify', -4117: 'distributed'}

def com_style(cell):
    def side(index):
        border = cell.Borders(index)
        if border.LineStyle == -4142: return '-'           # 선 없음
        if border.LineStyle == -4119: return 'd'           # 이중선
        return {2: 't', -4138: 'm', 4: 'k', 1: 'h'}.get(border.Weight, 't')     # 가는 선, 중간, 굵은 선, 아주 가는 선
    font = cell.Font
    return {'f': [font.Name, font.Size, 1 if font.Bold else 0],
            'b': side(7) + side(10) + side(8) + side(9),   # 왼쪽, 오른쪽, 위, 아래
            'a': [H_ALIGN.get(cell.HorizontalAlignment, ''), V_ALIGN.get(cell.VerticalAlignment, ''), 1 if cell.WrapText else 0, 1 if cell.ShrinkToFit else 0],
            'n': 'General' if cell.NumberFormat in ('G/표준', 'General') else cell.NumberFormat}     # 한글 엑셀은 '표준'을 G/표준 이라고 돌려준다

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
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 3        # 매크로 끔
    excel.EnableEvents = False
    styles, out = [], {}
    def sid(cell):
        s = com_style(cell)
        if s not in styles: styles.append(s)
        return styles.index(s)
    try:
        book = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
        for name, (title_row, first) in ROSTERS.items():
            ws, sheet = wb[name], book.Worksheets(name)
            last = max(c.column for row in ws.iter_rows(min_row=title_row, max_row=first - 1) for c in row if c.value is not None)
            title = re.match(r'=\s*"(.*?)\("\s*&\s*메뉴!E2\s*&\s*"\)"', ws.cell(title_row, 1).value).group(1)
            ps, pm = ws.page_setup, ws.page_margins
            head = []       # 제목 행부터 머리글 끝까지 모든 칸: [행, 열, 글자(없으면 None), 서식 번호]
            for r in range(title_row, first):
                for c in range(1, last + 1):
                    value = ws.cell(r, c).value
                    head.append([r, c, value if isinstance(value, str) and not value.startswith('=') else None, sid(sheet.Cells(r, c))])
            row_styles, row_heights, previous, previous_height = [], [], None, None
            for r in range(first, first + DATA_ROWS):
                ids = [sid(sheet.Cells(r, c)) for c in range(1, last + 1)]
                if ids != previous: row_styles.append([r, ids]); previous = ids
                height = round(sheet.Rows(r).RowHeight, 2)
                if height != previous_height: row_heights.append([r, height]); previous_height = height
            out[name] = {
                'titleRow': title_row, 'first': first, 'cols': last, 'title': title,
                'widths': widths(ws, last),
                'hiddenCols': sorted({i for d in ws.column_dimensions.values() if d.hidden for i in range(d.min, d.max + 1) if i <= last}),
                'heights': {r: ws.row_dimensions[r].height for r in range(1, first)},
                'merges': sorted(str(x) for x in ws.merged_cells.ranges if x.min_row >= title_row and x.max_row < first),
                'head': head, 'rowStyles': row_styles, 'rowHeights': row_heights,
                'data': [s for r, s in row_styles if r <= first + 1][-1],      # 둘째 자료 줄의 서식 (화면 미리보기의 숫자 표시 형식에 씀)
                'page': {'orientation': ps.orientation, 'paper': ps.paperSize, 'scale': ps.scale, 'hc': bool(ws.print_options.horizontalCentered),
                         'margins': [round(x, 4) for x in (pm.left, pm.right, pm.top, pm.bottom, pm.header, pm.footer)], 'titles': (ws.print_title_rows or '').replace('$', '')},
            }
            print(name, '| 자료 줄 서식이 바뀌는 행', [r for r, _ in row_styles], '| 높이가 바뀌는 행', row_heights)
        book.Close(False)
    finally:
        for other in list(excel.Workbooks): other.Close(False)
        excel.Quit()
    ws = wb['자료']
    data = {'widths': widths(ws, 33), 'merges': sorted(str(x) for x in ws.merged_cells.ranges),
            'row1': {c.coordinate: c.value for c in ws[1] if c.value is not None and c.column > 1}, 'row2': [c.value for c in ws[2][:33]], 'h1': ws.row_dimensions[1].height}
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
