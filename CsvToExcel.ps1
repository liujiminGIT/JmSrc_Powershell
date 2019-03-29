param(
    [parameter(mandatory=$true)][string]$batid,
    [parameter(mandatory=$true)][string]$xlsFormat,
    [parameter(mandatory=$true)][string]$xlsDelCol
)

#$batid="IFR080"
#$xlsFormat="13-14"
#$rootPath="C:\hzg\A-GPS\�J��\bat\"
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

 <# ���C�u�����Q�� #>
 [System.Reflection.Assembly]::LoadFrom($dllPath) | Out-Null

 <# Excel�̃t�H�[�}�b�g�ɓ]�L���āA�o�͐�ɖ��O��t���ĕۑ����� #>
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
       
        
        <# ��ڐ���2���̕�����ύX�i3��ڂ̏ꍇ�A03�ɂȂ�G10��̏ꍇ�A10�ɂȂ� #>
        $colStr= if($csvCol -lt 10){"0"+ [string]$csvCol} else {[string]$csvCol}
        
        <# �i�p�����[�^�jI/F����w�肵�����O���ڂ𔲂���@#>
        if($xlsDelCol.Contains($colStr)){ 
           continue
        }else{
           $col = $col + 1
        }
        
        <# �i�p�����[�^�j�Ώې��l��̏ꍇ�𐔎��^�ɕϊ����� #>
        $value = if($xlsFormat.Contains($colStr)){ [double]($csvRowData[$ci]) } else { [string]$csvRowData[$ci] }
        <# �@�l���̂��Œ�l���Z�b�g���� #>
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
        <# �l���Z�b�g����F2��ڂ͖@�l���̌Œ� #>
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