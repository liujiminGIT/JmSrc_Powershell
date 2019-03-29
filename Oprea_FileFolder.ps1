<#
STEP1: .wss ⇒ .zip Rename
STEP2: .zip ⇒ folder unzip
#>

# STEP1.wssファイルをzipにリネームする
Set-Location -Path 'C:\jm_work\7000_WK\abc'
Get-ChildItem -Recurse | Where-Object{$_.Extension -eq '.wss'} | Rename-Item -NewName {[io.path]::ChangeExtension($_.Name, "zip")}

# STEP2.ZIPファイル解凍して、元ファイル
$sh = new-object -com shell.application
$Zips = Get-ChildItem -Recurse | Where-Object{$_.Extension -eq '.zip'}
foreach($zip in $Zips)
{
    Write-Host $zip.FullName + "コピー開始します。"
    $zipPath = $zip.FullName.TrimEnd(".zip")
    New-Item -Path $zipPath -ItemType directory

    $targetfolder = $sh.namespace($zipPath)
    
    $zipFile = $sh.namespace($zip.FullName)

    $zipFile.Items() | ForEach-Object {
    $targetfolder.copyhere($_.Path, 0x14) 
    }

    Write-Host $zip.FullName + "コピー完了しました。"

}

<#
フォルダ内のファイルを一括リネーム（タイムスタンプ削除など）
#>
$Pt='D:\7000_作業WK\AGF14_0801_属性障害対応\20170801_060000'
$rp='_20170801 060000'
Get-ChildItem -path $pt | Rename-Item -NewName {$_.Name -replace $rp, ''}

<#
大きなテキストファイルを分割
#>
Set-Location C:\jm_work\7000_WK\biglog
$i=0; cat .\olap_server_bk20180510a.log -ReadCount 1000000 | % { $_ > test$i.txt;$i++ }



<#
SQLサーバーからエクスポートしたプロシージャを分割する、最初頭部分の行を削除してから、本スクリプトを実行
#>

set-location 'C:\jm_work\7000_WK\prc'

$script = Get-Content -Encoding 'UTF8' .\script.sql
#/****** Object:  StoredProcedure [dbo].[AddLogCnt]    Script Date: 2018/12/17 14:12:04 ******/
$key = '/****** Object:  StoredProcedure'

Write-host ([String]::Format("{0}",$key))

$path = ""
foreach($txt in $script)
{
    if($txt -like ([String]::Format('{0}*', $key)))
    {
       $path=[String]::Format(".\\sub_script\\{0}.sql", (($txt.Split("]")[1]).Replace(".", "").Replace("[", "")));
    }
    $txt >> $path
}
