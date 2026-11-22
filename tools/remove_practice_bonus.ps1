# 영양교사·유치원 양식에 숨겨져 있던 교육실습협력학교 가산점 흔적을 지운다.
#
#   - 총점(D9) 수식에서 실습학교 가산점(P35)을 더하는 부분을 뺀다: 총점 = 학교경력점 + 가산점
#   - 숨겨진 35~36행의 이름표('교육실습협력학교'), 소계 수식, 종별 목록 설정을 지운다
#   - 숨은 목록의 '실습교'(AB98), '전입가산점'(AK103) 을 지운다
#   - 35~36행 자체는 숨긴 채로 남긴다 -> 아래쪽(학교 확인, 추가 입력 항목 F48~F57)의 셀 주소가 바뀌지 않는다.
#
# 직접 실행하지 말고 tools/build_template.py 로 실행한다. (시트 보호를 잠시 떼었다가 다시 붙여 준다)
#   python tools/build_template.py <양식.xlsx> <양식.xlsx> 5 remove-practice
param(
    [Parameter(Mandatory)] [string] $Source,
    [Parameter(Mandatory)] [string] $Destination,
    [int] $CareerRows = 5      # 쓰지 않음 (build_template.py 가 늘 넘겨 주는 값)
)
$ErrorActionPreference = 'Stop'
$Source = (Resolve-Path $Source).Path
$Destination = [System.IO.Path]::GetFullPath($Destination)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
try {
    $wb = $xl.Workbooks.Open($Source, 0, $true)
    $ws = $wb.Worksheets.Item('내신조사서')
    if ($ws.ProtectContents) { throw '시트 보호가 걸려 있습니다. tools/build_template.py 로 실행하세요.' }
    if ($ws.Range('B47').Text -ne '추가 입력 항목') { throw "B47이 '추가 입력 항목'이 아닙니다: $($ws.Range('B47').Text)" }
    if (-not $ws.Rows(35).Hidden -or -not $ws.Rows(36).Hidden) { throw '35~36행이 숨겨져 있지 않습니다. 실습학교 칸이 숨겨진 양식에만 씁니다.' }
    if (([string]$ws.Range('B35').Text -replace '\s', '') -ne '교육실습협력학교') { throw "B35가 '교육실습협력학교'가 아닙니다: $($ws.Range('B35').Text)" }

    $total = [string]$ws.Range('D9').Formula
    if ($total -notmatch '^=IF\(ISERR\(Q12\+(P\d+)\),"",IF\(H5=AB99, ?Q12\+\1\+P35,Q12\+\1\)\)$') { throw "총점(D9) 수식이 예상과 다릅니다: $total" }
    $bonus = $Matches[1]
    $ws.Range('D9').Formula = "=IF(ISERR(Q12+$bonus),`"`",Q12+$bonus)"

    $ws.Range('B35:Q36').ClearContents() | Out-Null
    $ws.Range('C35:D36').Validation.Delete()
    foreach ($address in 'AB98', 'AK103') {
        $text = [string]$ws.Range($address).Text
        if ($text -notin '실습교', '전입가산점') { throw "$address 가 예상과 다릅니다: $text" }
        $ws.Range($address).ClearContents() | Out-Null
    }
    if ($ws.Range('B47').Text -ne '추가 입력 항목') { throw "'추가 입력 항목' 위치(B47)가 달라졌습니다." }

    $newTotal = [string]$ws.Range('D9').Formula
    $ws.Activate(); $ws.Range('B3').Select() | Out-Null
    $wb.SaveAs($Destination, 51)          # 51 = xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
    Write-Output "저장: $Destination (총점 수식: $newTotal)"
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
