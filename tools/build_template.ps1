# 2026.3.1.자 내신 조사서 양식 -> 2027.3.1.자 양식 변환 (엑셀을 직접 구동해 서식을 그대로 보존)
#
# 변경 내용
#   1) 실습교 보정 총점(O6/P6) 삭제, 총점(D9) = 학교경력점 + 가산점
#   2) 교육실습협력학교 가산점 입력칸(33~36행) 삭제 -> 일반 가산점 입력 행으로 전환
#      (행 번호를 유지하므로 아래쪽 '추가 입력 항목'(F48~F57) 주소가 바뀌지 않는다.
#       학교관리자취합용 파일이 이 주소를 그대로 읽는다.)
#   3) '작성 방법' 시트: 2027학년도 / 실습교 관련 항목 2개 삭제 / 날짜 예시 갱신
#   4) -CareerRows 6~8: 경력 사항 줄을 늘린 양식 (프로그램이 경력이 6줄 이상일 때만 사용)
#
# 직접 실행하지 말고 tools/build_template.py 로 실행한다.
# (내신조사서 시트에 암호 보호가 걸려 있어, build_template.py 가 보호를 잠시 떼었다가
#  변환 후 원래 보호 설정(같은 암호)을 그대로 다시 붙인다.)
param(
    [Parameter(Mandatory)] [string] $Source,
    [Parameter(Mandatory)] [string] $Destination,
    [ValidateRange(5, 8)] [int] $CareerRows = 5      # 경력 사항 줄 수 (기본 5줄, 최대 8줄)
)
$ErrorActionPreference = 'Stop'
$Source = (Resolve-Path $Source).Path
$Destination = [System.IO.Path]::GetFullPath($Destination)

$xlMedium = -4138; $xlThin = 2; $xlContinuous = 1
$xlEdgeLeft = 7; $xlEdgeTop = 8; $xlEdgeBottom = 9; $xlEdgeRight = 10
$xlPasteFormats = -4122; $xlPasteValidation = 6
$missing = [Type]::Missing

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Source, 0, $true)

    # ------------------------------------------------------------ 내신조사서
    $ws = $wb.Worksheets.Item('내신조사서')
    if ($ws.ProtectContents) { throw '시트 보호가 걸려 있습니다. tools/build_template.py 로 실행하세요.' }

    # (1) 실습교 보정 총점 삭제
    if ($ws.Range('O6').Text -notmatch '실습교') { throw "O6가 '실습교 보정 총점'이 아닙니다: $($ws.Range('O6').Text)" }
    $ws.Range('P6:Q6').UnMerge()
    $ws.Range('O6:Q6').ClearContents() | Out-Null
    $ws.Range('O6:Q6').Merge()
    $b = $ws.Range('O6:Q6').Borders($xlEdgeRight); $b.LineStyle = $xlContinuous; $b.Weight = $xlMedium
    $ws.Range('D9').Formula = '=IF(ISERR(Q12+P18),"",Q12+P18)'

    # (2) 교육실습협력학교 가산점 입력칸 삭제 -> 가산점 행으로 전환
    if ($ws.Range('B33').Text -notmatch '실습') { throw "B33이 교육실습협력학교 가산점 칸이 아닙니다." }
    foreach ($r in 'B17:B32', 'B33:B36', 'P18:Q32', 'P33:Q36') { $ws.Range($r).UnMerge() }
    $ws.Range('B33:Q36').ClearContents() | Out-Null
    $ws.Range('C32:O32').Copy() | Out-Null                       # 가산점 마지막 행(아래 굵은 테두리)
    $ws.Range('C36:O36').PasteSpecial($xlPasteFormats) | Out-Null
    $ws.Range('C31:O31').Copy() | Out-Null                       # 가산점 중간 행
    $ws.Range('C32:O35').PasteSpecial($xlPasteFormats) | Out-Null
    $ws.Range('C31:D31').Copy() | Out-Null                       # 종별 선택 목록
    $ws.Range('C32:D36').PasteSpecial($xlPasteValidation) | Out-Null
    $xl.CutCopyMode = 0
    $ws.Rows('32:36').RowHeight = $ws.Rows(31).RowHeight
    $ws.Range('B17:B36').Merge()
    $ws.Range('P18:Q36').Merge()
    $ws.Range('P18').Formula = '=IF(ISERR(SUM(O18:O36)),"",SUM(O18:O36))'
    foreach ($edge in @(
            @('B17:B36', $xlEdgeLeft, $xlMedium), @('B17:B36', $xlEdgeRight, $xlThin), @('B17:B36', $xlEdgeBottom, $xlMedium),
            @('P18:Q36', $xlEdgeLeft, $xlThin), @('P18:Q36', $xlEdgeRight, $xlMedium), @('P18:Q36', $xlEdgeBottom, $xlMedium))) {
        $b = $ws.Range($edge[0]).Borders($edge[1]); $b.LineStyle = $xlContinuous; $b.Weight = $edge[2]
    }
    $ws.Range('AB98:AB99').ClearContents() | Out-Null             # 숨김 목록의 실습교 항목
    if ($ws.Range('AK103').Text -eq '전입가산점') { $ws.Range('AK103').ClearContents() | Out-Null }

    # (2-1) 경력 사항 줄 늘리기 (기본 5줄: 12~16행)
    #       늘린 만큼 가산점 줄을 줄여 전체 행 수를 유지한다 -> 가산점 아래쪽 셀 주소는 그대로.
    $extra = $CareerRows - 5
    if ($extra -gt 0) {
        $lastCareer = 16 + $extra                                 # 늘어난 뒤 경력 마지막 행
        foreach ($r in 'B11:B16', 'C12:D16', 'Q12:Q16') { $ws.Range($r).UnMerge() }
        $ws.Rows("17:$lastCareer").Insert() | Out-Null            # 아래(가산점~추가 입력 항목)가 $extra 행 밀림
        $ws.Range('B16:Q16').Copy() | Out-Null                    # 경력 마지막 행(아래 굵은 테두리)
        $ws.Range("B${lastCareer}:Q${lastCareer}").PasteSpecial($xlPasteFormats) | Out-Null
        $ws.Range('B15:Q15').Copy() | Out-Null                    # 경력 중간 행
        $ws.Range("B16:Q$($lastCareer - 1)").PasteSpecial($xlPasteFormats) | Out-Null
        $ws.Range('S16:Y16').Copy($ws.Range("S17:Y${lastCareer}")) | Out-Null   # 숨김 계산 열
        $xl.CutCopyMode = 0
        $ws.Rows("17:$lastCareer").RowHeight = $ws.Rows(16).RowHeight
        $ws.Range("B11:B${lastCareer}").Merge()
        $ws.Range("C12:D${lastCareer}").Merge()
        $ws.Range("Q12:Q${lastCareer}").Merge()
        $ws.Range('Q12').Formula = "=IF(ISERR(SUM(P12:P${lastCareer})),`"`",SUM(P12:P${lastCareer}))"

        # 가산점 줄(마지막 줄 바로 위)을 같은 수만큼 삭제 -> 아래쪽이 원래 위치로 돌아옴
        $ws.Rows("36:$(35 + $extra)").Delete() | Out-Null
        $bonusTop = 18 + $extra
        $ws.Range("P${bonusTop}").Formula = "=IF(ISERR(SUM(O${bonusTop}:O36)),`"`",SUM(O${bonusTop}:O36))"
        $ws.Range('D9').Formula = "=IF(ISERR(Q12+P${bonusTop}),`"`",Q12+P${bonusTop})"
        if ($ws.Range('B47').Text -ne '추가 입력 항목') { throw "줄을 늘린 뒤 '추가 입력 항목' 위치(B47)가 달라졌습니다: $($ws.Range('B47').Text)" }
    }

    # ------------------------------------------------------------ 작성 방법
    $g = $wb.Worksheets.Item('작성 방법')
    $g.Range('B1').Value2 = '2027학년도 관내 전보 내신 조사서 작성 방법'

    function Find-GuideRow([string] $label) {
        foreach ($r in 4..60) { if ([string]$g.Cells.Item($r, 3).Value2 -eq $label) { return $r } }
        throw "작성 방법 시트에서 '$label' 항목을 찾지 못했습니다."
    }
    foreach ($label in '교육실습 협력학교 가산점', '실습교 보정총점') {
        $row = Find-GuideRow $label
        for ($i = $g.Shapes.Count; $i -ge 1; $i--) {              # 그 행에 놓인 작성 예시 그림도 함께 삭제
            if ($g.Shapes.Item($i).TopLeftCell.Row -eq $row) { $g.Shapes.Item($i).Delete() }
        }
        $g.Rows($row).Delete() | Out-Null
    }
    $last = 4
    foreach ($r in 4..60) { if ([string]$g.Cells.Item($r, 3).Value2) { $last = $r } }
    foreach ($r in 4..$last) { $g.Cells.Item($r, 2).Value2 = $r - 4 }   # 순 번호 다시 매기기

    $r = Find-GuideRow '경력사항 근무기간'
    $g.Cells.Item($r, 5).Value2 = "2023.03.01~2027.02.28`n2022.03.01~2023.02.28"
    $g.Cells.Item($r, 6).Value2 = "2023.03.01.-2027.02.28.`n2023.03.01.~2027.02.28.`n23.03.01~27.02.28`n23.3.1~24.2.29`n2023년3월1일~2027년2월28일 등"
    $r = Find-GuideRow '가산점 기간,  연월일'
    $g.Cells.Item($r, 5).Value2 = "2025.03.01~2027.02.28`n2025.08.29 등`n"
    $g.Cells.Item($r, 6).Value2 = "2025.03.01.-2027.02.28.`n2025.03.01.~2027.02.28.`n25.03.01~27.02.28`n25.3.1~26.2.28`n2025년3월1일~2027년2월28일`n2025.8.29, 2025년8월29일 등"
    $r = Find-GuideRow '가산점 내용'
    $cell = $g.Cells.Item($r, 5)
    $at = ([string]$cell.Value2).IndexOf('자녀가 2명인 교사')
    if ($at -ge 0) { $cell.Characters($at + 1 + '자녀가 2명인 '.Length, 2).Text = '교원' }   # 글꼴 서식 유지

    $g.Activate(); $g.Range('B1').Select() | Out-Null
    $ws.Activate(); $ws.Range('B3').Select() | Out-Null
    $wb.SaveAs($Destination, 51)          # 51 = xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
    Write-Output "저장: $Destination"
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
