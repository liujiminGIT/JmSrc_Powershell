<#

#>

# M事得意先別予算データXlsxファイル
$xlsxPath = 'C:\jm_work\2500_AJC03\2019年度切替作業WK\002_M事業得意先別'
$logFile = 'C:\jm_work\2500_AJC03\2019年度切替作業WK\002_M事業得意先別\ALL.log'
Set-Location $xlsxPath


$excel = New-Object -ComObject excel.application
$excel.visible = $false
$excel.DisplayAlerts = $false

$files = Get-ChildItem -Path $xlsxPath -Recurse | Where-Object{$_.Extension -eq '.xlsx'} 

$row_Jigyo_start = 5
$row_head_start  = 7
$row_data_a      = 8
$row_data_b      = 223

$blockRows = 221
foreach($f in $files)
{
  $txtFile = ($xlsxPath + "\" + $f.BaseName + ".txt")
  if((Test-Path -Path $txtFile) -eq $true)
  {
      Remove-Item -Path ($txtFile) -Force
  }
  New-Item -Path ($txtFile) -ItemType File -Force
  $book = $excel.Workbooks.Open($f.FullName)
  $sheets = $book.Worksheets
  foreach($s in $sheets)
  {
    if($s.Visible)
    {
      $log = ([String]::Format("【{0:yyyyMMdd_HHmmss}】{1}-{2}処理開始", (Get-Date), $f.BaseName ,$s.Name))
      $log >> $logFile
      Write-Host $log 
      
      for($l = 0; $l -le 100; $l++)
      {
        $row_Jigyo = $row_Jigyo_start + $l * $blockRows
        $row_head  = $row_head_start  + $l * $blockRows
        $row_data_s  = $row_data_a    + $l * $blockRows
        $row_data_e  = $row_data_b    + $l * $blockRows
        
        $customer = "得意先"
        $kanjyo = "勘定科目"
        $jigyo = $s.Cells.Item($row_Jigyo, 2).Value()

        if([String]::IsNullOrEmpty($jigyo))
        {
          break
        }

        $recode = [String]::Format("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}`t{9}`t{10}`t{11}`t{12}`t{13}`t{14}`t{15}`t{16}`t{17}`t{18}`t{19}"
        ,"ファイル名称"
        ,"シート名称"
        ,"事業L4"
        ,"得意先"
        ,"勘定科目"
        ,$s.Cells.Item($row_head, 4).Value()
        ,$s.Cells.Item($row_head, 5).Value()
        ,$s.Cells.Item($row_head, 6).Value()
        ,$s.Cells.Item($row_head, 7).Value()
        ,$s.Cells.Item($row_head, 8).Value()
        ,$s.Cells.Item($row_head, 9).Value()
        ,$s.Cells.Item($row_head, 10).Value()
        ,$s.Cells.Item($row_head, 11).Value()
        ,$s.Cells.Item($row_head, 12).Value()
        ,$s.Cells.Item($row_head, 13).Value()
        ,$s.Cells.Item($row_head, 14).Value()
        ,$s.Cells.Item($row_head, 15).Value()
        ,$s.Cells.Item($row_head, 16).Value()
        ,$s.Cells.Item($row_head, 17).Value()
        ,$s.Cells.Item($row_head, 18).Value() )

        $recode >> $txtFile 


        for($r = $row_data_s; $r -le $row_data_e; $r++)
        {
          $col_b = $s.Cells.Item($r, 2).Value()
          $kanjyo = $s.Cells.Item($r, 3).Value()
          if($kanjyo -eq "売上高")
          {
            $customer=$col_b
          }

          $recode = [String]::Format("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}`t{9}`t{10}`t{11}`t{12}`t{13}`t{14}`t{15}`t{16}`t{17}`t{18}`t{19}"
          ,$f.BaseName
          ,$s.Name
          ,$jigyo
          ,$customer
          ,$kanjyo
          ,$s.Cells.Item($r, 4).Value()
          ,$s.Cells.Item($r, 5).Value()
          ,$s.Cells.Item($r, 6).Value()
          ,$s.Cells.Item($r, 7).Value()
          ,$s.Cells.Item($r, 8).Value()
          ,$s.Cells.Item($r, 9).Value()
          ,$s.Cells.Item($r, 10).Value()
          ,$s.Cells.Item($r, 11).Value()
          ,$s.Cells.Item($r, 12).Value()
          ,$s.Cells.Item($r, 13).Value()
          ,$s.Cells.Item($r, 14).Value()
          ,$s.Cells.Item($r, 15).Value()
          ,$s.Cells.Item($r, 16).Value()
          ,$s.Cells.Item($r, 17).Value()
          ,$s.Cells.Item($r, 18).Value() )
          $recode >> $txtFile 
        }
      }
      $log = Write-Host ([String]::Format("【{0:yyyyMMdd_HHmmss}】{1}-{2}処理終了", (Get-Date), $f.BaseName ,$s.Name))
      $log >> $logFile
      Write-Host $log 
    }
  }
  $book.Close()
}
$excel.Quit()
[GC]::Collect()

# STEP2ファイルマージ
Get-ChildItem $xlsxPath -Recurse -File -Filter "*.txt" | Get-Content | Add-Content ($xlsxPath+"\Total.dat")
