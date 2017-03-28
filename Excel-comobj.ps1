$excel = New-Object -ComObject Excel.Application

$excel.Visible = $true

$workbook = $excel.Workbooks.Add()

$sheet = $workbook.ActiveSheet

$counter = 0

Get-process | ForEach-Object {

    $counter++

    $sheet.cells.Item($counter,1) = $_.Name

    $sheet.cells.Item($counter,2) = $_.DisplayName

    $sheet.cells.Item($counter,3) = $_.Status

}