"""엑셀 양식 파일을 transfer_application.html 안에 넣는다.

브라우저는 HTML 파일을 더블클릭으로 직접 열었을 때(주소가 file:// 로 시작) 옆에 있는 엑셀 파일을
읽지 못하게 막는다. 그래서 양식 파일 내용을 HTML 안에 함께 넣어 두고, 프로그램은 그 내용을 쓴다.

template.xlsx 나 template_career6~8.xlsx 를 고친 뒤에는 이 스크립트를 다시 실행해야 화면에 반영된다.

    python tools/embed_templates.py
"""
import base64
import pathlib
import sys

sys.stdout.reconfigure(encoding='utf-8')
ROOT = pathlib.Path(__file__).resolve().parent.parent
HTML = ROOT / 'transfer_application.html'
TEMPLATES = [base.replace('.xlsx', suffix + '.xlsx')
             for base in ('template.xlsx', 'template_nutrition.xlsx', 'template_kindergarten.xlsx')
             for suffix in ('', '_career6', '_career7', '_career8')]
BEGIN = '    <script id="embeddedTemplates">\n'
END = '    </script><!-- /embeddedTemplates -->\n'

src = HTML.read_text(encoding='utf-8')
newline = '\r\n' if '\r\n' in src else '\n'
src = src.replace('\r\n', '\n')

lines = ['        // tools/embed_templates.py 가 만든 내용입니다. 직접 고치지 마세요.',
         '        const EMBEDDED_TEMPLATES = {']
for name in TEMPLATES:
    data = base64.b64encode((ROOT / name).read_bytes()).decode('ascii')
    lines.append(f"            '{name}': '{data}',")
    print(f'{name}: {len(data):,}자')
lines.append('        };')
block = BEGIN + '\n'.join(lines) + '\n' + END

if BEGIN in src:
    start = src.index(BEGIN)
    end = src.index(END, start) + len(END)
    src = src[:start] + block + src[end:]
else:
    marker = '    <script>\n        const gradeData'
    assert src.count(marker) == 1, '넣을 위치를 찾지 못했습니다'
    src = src.replace(marker, block + marker)

HTML.write_text(src.replace('\n', newline), encoding='utf-8', newline='')
print('transfer_application.html 에 양식을 넣었습니다.')
