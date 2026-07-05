$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open((Resolve-Path "Teacher_Mapping.xlsx").Path)
$ws = $wb.Sheets.Item(1)
for ($r=1; $r -le 10; $r++) {
    $rowStr = ""
    for ($c=1; $c -le 5; $c++) {
        $cell = $ws.Cells.Item($r, $c).Text
        $rowStr += "$cell | "
    }
    Write-Host $rowStr
}
$wb.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
