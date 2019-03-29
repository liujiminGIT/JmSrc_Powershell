param(
    [parameter(mandatory=$true)][string]$batid,
    [parameter(mandatory=$true)][string]$xlsFormat,
    [parameter(mandatory=$true)][string]$xlsDelCol
)

#$batid="IFR080"
#$xlsFormat="13-14"
#$rootPath="C:\hzg\A-GPS\開発\bat\"
$rootPath="J:\Cognos\TM1\AGPS\Dat\RECV\"
#$filePath=$rootPath + $batid + "\"
$filePath=$rootPath
$templatePath=$filePath + "TMP_" + $batid + ".xltx"
$csvPath=$filePath + $batid + ".csv"
$xlsFilePath=$filePath + $batid + ".xlsx"
$dllPath= "J:\Cognos\TM1\AGPS\Scripts\lib\EPPlus.dll"
$companyShortName=""

$rc=0
try{

 <# ライブラリ参照 #>
 [System.Reflection.Assembly]::LoadFrom($dllPath) | Out-Null

 <# Excelのフォーマットに転記して、出力先に名前を付けて保存する #>
 $pkg = New-Object OfficeOpenXml.ExcelPackage -ArgumentList $templatePath
 $sht = $pkg.Workbook.Worksheets[1]

 $csvRows= (Get-Content -Path $csvPath -Encoding UTF8) -as [string[]]

 for($ri=0; $ri -lt $csvRows.length; $ri++) {

    $csvRowData = ($csvRows[$ri]) -split "`t"
   # $csvRowData = $csvRowData1.split(",")
    $col=0
    for($ci=0; $ci -lt $csvRowData.length; $ci++) {

        $row = $ri + 2
        $csvCol= $ci + 1
       
        
        <# 列目数を2桁の文字列変更（3列目の場合、03になる；10列の場合、10になる #>
        $colStr= if($csvCol -lt 10){"0"+ [string]$csvCol} else {[string]$csvCol}
        
        <# （パラメータ）I/Fから指定した除外項目を抜ける　#>
        if($xlsDelCol.Contains($colStr)){ 
           continue
        }else{
           $col = $col + 1
        }
        
        <# （パラメータ）対象数値列の場合を数字型に変換する #>
        $value = if($xlsFormat.Contains($colStr)){ [double]($csvRowData[$ci]) } else { [string]$csvRowData[$ci] }
        <# 法人略称を固定値をセットする #>
        if($col -eq 1){
           $companyShortName=""
           switch ($value) {
             "0001" {
                 $companyShortName="A"
             }
             "0002" {
                 $companyShortName="B"
             }
             "0003" {
                 $companyShortName="C"
             }
             "0004" {
                 $companyShortName="D"
             }
             "0005" {
                 $companyShortName="E"
             }
             "0006" {
                 $companyShortName="F"
             }
             "0007" {
                 $companyShortName="G"
             }
             "0007" {
                 $companyShortName="H"
             }
           }
        }
        <# 値をセットする：2列目は法人略称固定 #>
        if($col -eq 2){
           $sht.Cells.Item($row, $col).Value =  $companyShortName
        } else {
           $sht.Cells.Item($row, $col).Value =  $value
        }
    }
}

 $pkg.SaveAs($xlsFilePath)

}catch [Exception]{
#  Write-host $error
 $rc = 1
}finally{
 $pkg.Dispose()
}
exit $rc