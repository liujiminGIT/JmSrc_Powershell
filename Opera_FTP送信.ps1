Param(
    [parameter(mandatory=$true)][String]$ENVTYPE="DEV"
)

$tmstp = (Get-Date).ToString("yyyyMMdd_HHmmss_")
$log_path = "J:\bat\log\AHS04\AHS04_JDX2DSM_DATAHUB_$tmstp.log"

Start-Transcript -Path $log_path -Append
Import-Module "I:\bat\AHS04\AHS04_COM_UTILITY.psm1"


#STEP1 IF出力ETL起動
$prj = "AHS04_データ出力_Jedox2DataHub_一括"
$job = "Jb_Pl_Powershell起動_IF送信"

Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("ETL処理開始 $prj _ $job"))
$etlResult = runETL -project $prj -job $job
if(($etlResult[0] -ne 10) -and ($etlResult[0] -ne 20))
{
    $errmsg = [String]::Format("Job：{0} `r`nExitCode：{1} https://knowledgebase.jedox.com/knowledgebase/command-line-client/ `r`n終了時刻{2}",$job1, $runETL.ExitCode, $runETL.ExitTime)
    SendMail -Title ($prj1+"失敗") -Text ($errmsg) -Priority "H"
    Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("ETL処理失敗 $errmsg"))
    Exit
}
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("ETL処理終了 $prj _ $job"))


#STEP2 HULFT送信バッチ起動
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("HULFT送信開始"))
<#
$hulftid =@(
"DRS1001F", #T_SALES_SO_ACT_D.txt
"DRS1002F", #M_ACCOUNT.txt
"DRS1003F", #M_CUSTOMER.txt
"DRS1004F", #M_DEL_DEST.txt
"DRS1005F", #M_PROD.txt
"DRS1006F", #T_SALES_ACT_D.txt
"DRS1006F"  #DATAHUB_END.txt
)
$errMsg = @()
foreach($id in $hulftid)
{
    $hulftResult = execHULFT -hulftid $id
    if($hulftResult[0] -ne 0)
    {
        $errMsg += $hulftResult[1]
    }
    Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("HULFTID : $id  結果 : $hulftResult"))
}

if($errMsg.length -ne 0)
{
    $txt = $errMsg -join "`r`n"
    SendMail -Title ("HULFT送信失敗") -Text ($txt) -Priority "H"
    Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("HULFT送信異常終了  $txt"))
    Exit
}

#>
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("システムテストのため、当面HULFT送信しない"))
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("HULFT送信終了"))

# FTP接続に必要な情報を設定
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("DrSumIF FTP 送信開始"))
$drsumserver   = 'FTPServer';
$drsumuser     = 'username';
$drsumpass     = 'password';
# FTP接続用のURL
$drsumUrl = "ftp://$drsumserver/";

# 接続
$drsumwebClient = New-Object System.Net.WebClient;
$drsumwebClient.Credentials = New-Object System.Net.NetworkCredential($drsumuser,$drsumpass);
$drsumwebClient.BaseAddress = $drsumUrl;

$srcPath = "J:\DAT\SEND\AHS04\DATAHUB\"
$files = Get-ChildItem -Path $srcPath
foreach($f in $files)
{
    $fname = $f.Name
    $localFilePath = $f.FullName
    $drsumserverFilePath = "\$ENVTYPE\SEND\DATAHUB\$fname";
    # アップロード
    $drsumwebClient.UploadFile($drsumserverFilePath , $localFilePath);
    Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("FTP送信　$localFilePath⇒$drsumserverFilePath"))
}
$drsumwebClient.Dispose(); 
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("DrSumIF FTP 送信完了"))

#IF 退避処理
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("DATAHUB　IF退避開始"))
$tgtRootPath = "J:\DAT\SEND\BK_SEND\AHS04\DATAHUB\"
$tgtPath = ($tgtRootPath + $tmstp + "$ENVTYPE")

New-Item -Path $tgtPath -ItemType Directory
if(Test-Path -Path $tgtPath)
{
    Move-Item -Path ($srcPath + "*.*") -Destination $tgtPath
}
Write-Output ((Get-Date).ToString("yyyy/MM/dd HH:mm:ss ") +("DATAHUB　IF退避終了"))

Stop-Transcript
Exit



