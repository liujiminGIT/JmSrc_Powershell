
<#
Jedoxのリポジトリをプロジェクト毎に分割
#>>

# 解析対象のETLリポジトリパス
$path = 'C:\jm_work\7000_WK\aft09_etl'
# リポジトリファイル名称
$rep = 'repository_fallback.xml'
# 個別のETLを格納する対象フォルダ
$subpath = '\sub\'
<#
★★★この以下は修正しなくでよい★★★
#>
$sub=$path +$subpath
$file = $path + "\\$rep"

if((Test-Path $sub) -eq $true)
{
    Remove-Item -Path $sub -Force -Recurse
}

New-Item -Path $sub -itemType Directory

# リポジトリをxmlとして読み込み
$rep = [xml](Get-Content -Encoding 'UTF8' $file)

forEach($pj in $rep.configs.projects.ChildNodes)
{
    # プロジェクト毎のXMLファイルを出力する
    ([System.Xml.XmlDocument]$pj.OuterXml).Save($sub+$pj.name+'.xml')
    ($pj.name + "`t" + $pj.olapId) >> ($path + "\\list.log")
}



<#
JedoxのETLのリポジトリ操作（Project毎の説明、Job、説明）を抽出
#>>

$rep = [xml](Get-Content -Encoding 'UTF8' .\repository_fallback.xml)
forEach($pj in $rep.configs.projects.ChildNodes)
{
    forEach($jb in $pj.jobs.ChildNodes)
    {
        
        $pname = $pj.name
        $pcmt = if($pj.headers.header.comment.'#cdata-section' -eq $null){ "" } else {$pj.headers.header.comment.'#cdata-section'}
        $jname = $jb.name
        $jcmt = if ($jb.comment.'#cdata-section' -eq $null){ "" } else {$jb.comment.'#cdata-section'}

        $r =($pname + "`t" + $pcmt + "`t" +$jname+ "`t" +$jcmt).Replace("`r", " ").Replace("`n", " ")
        Write-Host $r
        $r | Out-File ("a.txt") -Encoding UTF8 -Append

        
    }
}