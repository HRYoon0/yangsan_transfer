# 내신 조사서 양식의 경력 사항 줄만 늘린다 (영양교사·유치원 양식용. 다른 내용은 바꾸지 않는다)
#
#   - 경력 사항(기본 5줄: 12~16행)을 6~8줄로 늘리고, 늘린 만큼 가산점 줄을 줄여 전체 행 수를 유지한다.
#     -> 가산점 아래쪽(학교 확인, 추가 입력 항목 F48~F57)의 셀 주소가 바뀌지 않는다.
#   - -LastBonusRow: 가산점 마지막 행. 영양교사·유치원 양식은 34행(35~36행은 숨겨진 빈 줄. 예전 교육실습협력학교 칸으로, 내용은 지웠다).
#
# 직접 실행하지 말고 tools/build_template.py 로 실행한다. (시트 보호를 잠시 떼었다가 다시 붙여 준다)
#   python tools/build_template.py <원본.xlsx> <결과.xlsx> <줄 수 6~8> career-only
param(
    [Parameter(Mandatory)] [string] $Source,
    [Parameter(Mandatory)] [string] $Destination,
    [ValidateRange(6, 8)] [int] $CareerRows = 6,
    [int] $LastBonusRow = 34
)
$ErrorActionPreference = 'Stop'
$Source = (Resolve-Path $Source).Path
$Destination = [System.IO.Path]::GetFullPath($Destination)
$xlPasteFormats = -4122

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Source, 0, $true)
    $ws = $wb.Worksheets.Item('내신조사서')
    if ($ws.ProtectContents) { throw '시트 보호가 걸려 있습니다. tools/build_template.py 로 실행하세요.' }
    if ($ws.Range('B11').Text -notmatch '경력') { throw "B11이 '경력사항'이 아닙니다: $($ws.Range('B11').Text)" }
    if ($ws.Range('B47').Text -ne '추가 입력 항목') { throw "B47이 '추가 입력 항목'이 아닙니다: $($ws.Range('B47').Text)" }
    $bonusFormula = [string]$ws.Range('P18').Formula
    if ($bonusFormula -notmatch "SUM\(O18:O$LastBonusRow\)") { throw "가산점 소계(P18) 범위가 O18:O$LastBonusRow 가 아닙니다: $bonusFormula" }

    $extra = $CareerRows - 5
    $lastCareer = 16 + $extra                                     # 늘어난 뒤 경력 마지막 행
    foreach ($r in 'B11:B16', 'C12:D16', 'Q12:Q16') { $ws.Range($r).UnMerge() }
    $ws.Rows("17:$lastCareer").Insert() | Out-Null                # 아래(가산점~추가 입력 항목)가 $extra 행 밀림
    $ws.Range('B16:Q16').Copy() | Out-Null                        # 경력 마지막 행(아래 굵은 테두리)
    $ws.Range("B${lastCareer}:Q${lastCareer}").PasteSpecial($xlPasteFormats) | Out-Null
    $ws.Range('B15:Q15').Copy() | Out-Null                        # 경력 중간 행
    $ws.Range("B16:Q$($lastCareer - 1)").PasteSpecial($xlPasteFormats) | Out-Null
    $ws.Range('S16:Y16').Copy($ws.Range("S17:Y${lastCareer}")) | Out-Null   # 숨김 계산 열
    $xl.CutCopyMode = 0
    $ws.Rows("17:$lastCareer").RowHeight = $ws.Rows(16).RowHeight
    $ws.Rows("17:$lastCareer").Hidden = $false
    $ws.Range("B11:B${lastCareer}").Merge()
    $ws.Range("C12:D${lastCareer}").Merge()
    $ws.Range("Q12:Q${lastCareer}").Merge()
    $ws.Range('Q12').Formula = "=IF(ISERR(SUM(P12:P${lastCareer})),`"`",SUM(P12:P${lastCareer}))"

    # 가산점 줄(마지막 줄 바로 위)을 같은 수만큼 삭제 -> 아래쪽이 원래 위치로 돌아옴
    $ws.Rows("${LastBonusRow}:$($LastBonusRow - 1 + $extra)").Delete() | Out-Null
    $bonusTop = 18 + $extra
    $ws.Range("P${bonusTop}").Formula = "=IF(ISERR(SUM(O${bonusTop}:O${LastBonusRow})),`"`",SUM(O${bonusTop}:O${LastBonusRow}))"
    if ($ws.Range('B47').Text -ne '추가 입력 항목') { throw "줄을 늘린 뒤 '추가 입력 항목' 위치(B47)가 달라졌습니다: $($ws.Range('B47').Text)" }
    $total = [string]$ws.Range('D9').Formula                      # 총점 수식은 엑셀이 행 이동에 맞춰 고쳐 준다
    if ($total -notmatch "Q12\+P${bonusTop}\b") { throw "총점(D9) 수식이 가산점 소계(P${bonusTop})를 가리키지 않습니다: $total" }

    $ws.Activate(); $ws.Range('B3').Select() | Out-Null
    $wb.SaveAs($Destination, 51)          # 51 = xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
    Write-Output "저장: $Destination (경력 ${CareerRows}줄, 가산점 ${bonusTop}~${LastBonusRow}행, 총점 수식: $total)"
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
