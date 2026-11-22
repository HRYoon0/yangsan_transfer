"""2026.3.1.자 내신 조사서 양식을 2027.3.1.자 양식으로 변환한다.

    python tools/build_template.py <원본.xlsx> <결과.xlsx> [경력 줄 수 5~8]

경력 줄 수를 6~8로 주면 경력 사항 줄을 늘린 양식을 만든다(가산점 줄이 그만큼 줄고 전체 행 수는 같다).
프로그램은 template.xlsx(5줄)를 기본으로 쓰고, 경력이 6줄 이상일 때만 template_career6~8.xlsx 를 쓴다.

영양교사·유치원 양식은 내용은 그대로 두고 경력 줄만 늘린다(끝에 career-only 를 붙인다).

    python tools/build_template.py template_nutrition.xlsx template_nutrition_career6.xlsx 6 career-only
    python tools/build_template.py template_kindergarten.xlsx template_kindergarten_career6.xlsx 6 career-only

영양교사·유치원 양식에 숨겨져 있던 교육실습협력학교 가산점 흔적을 지울 때는 끝에 remove-practice 를 붙인다(제자리에서 고침).

    python tools/build_template.py template_nutrition.xlsx template_nutrition.xlsx 5 remove-practice

실제 편집은 build_template.ps1 이 엑셀을 구동해서 한다(서식·그림·인쇄 설정 보존).
'내신조사서' 시트에는 암호 보호가 걸려 있어 엑셀이 편집을 거부하므로,
  1) 임시 사본에서 시트 보호 설정만 잠시 떼어 내고
  2) 엑셀로 변환한 뒤
  3) 결과 파일에 원래 보호 설정(같은 암호)을 그대로 다시 붙인다.
암호 자체는 알 필요도 없고 바뀌지도 않는다.
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

SHEET = 'xl/worksheets/sheet1.xml'          # 내신조사서
PROTECTION = re.compile(r'<sheetProtection\b[^>]*/>')


def rewrite(src, dst, edit):
    """src 를 dst 로 복사하면서 SHEET 의 XML 만 edit(xml) 결과로 바꾼다."""
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == SHEET:
                data = edit(data.decode('utf-8')).encode('utf-8')
            zout.writestr(info, data)


def main(source, destination, career_rows=5, career_only=False, script=None):
    here = os.path.dirname(os.path.abspath(__file__))
    work = tempfile.mkdtemp(prefix='naesin_template_')
    try:
        found = []

        def strip(xml):
            found.extend(PROTECTION.findall(xml))
            return PROTECTION.sub('', xml)

        unprotected = os.path.join(work, 'source_unprotected.xlsx')
        converted = os.path.join(work, 'converted.xlsx')
        rewrite(source, unprotected, strip)

        # career_only: 영양교사·유치원 양식처럼 경력 줄만 늘리는 경우 (초등 양식 전용 변환은 하지 않음)
        script = script or ('extend_career_rows.ps1' if career_only else 'build_template.ps1')
        subprocess.run(['pwsh', '-NoProfile', '-File', os.path.join(here, script),
                        '-Source', unprotected, '-Destination', converted,
                        '-CareerRows', str(career_rows)], check=True, timeout=300)

        def restore(xml):
            if not found:
                return xml
            assert not PROTECTION.search(xml)
            # sheetProtection 은 sheetData(및 sheetCalcPr) 바로 뒤에 와야 한다
            m = re.search(r'</sheetData>(\s*<sheetCalcPr\b[^>]*/>)?', xml)
            return xml[:m.end()] + found[0] + xml[m.end():]

        rewrite(converted, destination, restore)
        print('저장:', destination, '(시트 보호 유지)' if found else '(시트 보호 없음)')
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 5,
         career_only=len(sys.argv) > 4 and sys.argv[4] == 'career-only',
         script='remove_practice_bonus.ps1' if len(sys.argv) > 4 and sys.argv[4] == 'remove-practice' else None)
