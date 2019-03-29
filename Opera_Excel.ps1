
<#
Txtの内容をExcelへ転記
#>


Set-Location 'C:\jm_work\7000_WK\VT_MONTHLY_JIPROS\'

$xls='C:\jm_work\7000_WK\VT_MONTHLY_JIPROS\a.xlsx'
$txt = 'C:\jm_work\7000_WK\VT_MONTHLY_JIPROS\VT_MONTHLY_JIPROS.sql'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$book1 = $excel.Workbooks.open($xls)
$book1_s1 = $excel.Worksheets.Item(1)

$row = 2
$col = 2

 Get-Content $txt | ForEach-Object{
    $line = $_
    if(($line.Trim()) -ccontains 'UNION ALL')
    {
        $row = 2
        $col = $col + 1
    }
    else
    {
        $book1_s1.Cells.Item($row, $col)=$line
        $row = $row + 1
    }
}
$book1.SaveAs('C:\jm_work\7000_WK\VT_MONTHLY_JIPROS\b.xlsx')
$book1.close()

#$book1_s1 = $null
#$book =$null
#$excel = $null

[gc]::Collect()



<#########################################################################>

<#
フォルダ内のExcelのBookとSheet一覧取得
#>

Set-Location 'C:\jm_work\7000_WK\回収用_部・G確定版'

$excel = New-Object -ComObject excel.application
$excel.visible = $false
$excel.DisplayAlerts = $false
$tab = "`t"
$listFile = 'SYS_組織グループ一覧.txt'
(Date)|Set-Content $listFile -Encoding UTF8

$files = Get-ChildItem -Recurse | Where-Object{$_.Extension -eq '.xlsx'} 

foreach($f in $files)
{
  #Write-Host($f.FullName)
  
  $book = $excel.Workbooks.Open($f.FullName)
  $sheets = $book.Worksheets
  foreach($s in $sheets)
  {
    if($s.Visible)
    {
      ($f.Name + $tab + $s.Name) | Add-Content $listFile -Encoding UTF8 -Force
    }
  }
  $book.Close()
}
$excel.Quit()



<#########################################################################>