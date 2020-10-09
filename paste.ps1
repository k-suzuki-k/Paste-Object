$excel = New-Object -ComObject Excel.Application
$src = "C:\Users\xxxx\Desktop\ps1\graph.xlsm"
$book  = $excel.Workbooks.Open($src)

$files = Resolve-Path ".\files\*.xlsx"
$row = 1
foreach ($file in $files) {
    $rng = "A${row}"
    $excel.Run("pasteTo", $file.ToString(), "Sheet1", $rng) 
    $row += 13
}

$book.Save()
$book.Close()
$excel.Quit()

[void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
