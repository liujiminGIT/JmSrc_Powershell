<#
各種共通関数のまとめ
#>


<#
メール送信関数
#>
function SendMail
{
    param(
          [parameter(position=0)]
          $Title,
          [parameter(position=1)]
          $Text,
          [parameter(position=2)]
          $Priority
         )

    $From="sekimin_ryu@Mail.com"
    $To="sekimin_ryu@Mail.com"
    
    $Subject=$env:computername + $Title
    
    $body=$Text

    # 送信メールサーバーの設定
    $SMTPServer="xxx.xxx.xxx.xxx"
    $Port="25"
    $User="sekimin_ryu@Mail.com"
    $Password=
    $SMTPClient=New-Object Net.Mail.SmtpClient($SMTPServer,$Port)
    # SSL暗号化通信しない $false
    $SMTPClient.EnableSsl=$false
    $SMTPClient.Credentials=New-Object Net.NetworkCredential($User,$Password)

    # メールメッセージの作成
    $MailMassage=New-Object Net.Mail.MailMessage($From,$To,$Subject,$body)
    if($Priority -eq "H")
    {
        $MailMassage.Priority=[System.Net.Mail.MailPriority]::High
    }
    elseif($Priority -eq "L")
    {
        $MailMassage.Priority=[System.Net.Mail.MailPriority]::Low
    }

    
    # ファイルから添付ファイルを作成
    # $Attachment=New-Object Net.Mail.Attachment($File)
    # メールメッセージに添付
    #$MailMassage.Attachments.Add($Attachment)
    # メールメッセージを送信
    $SMTPClient.Send($MailMassage)
}

<#
JedoxETL処理実行用
#>

function runJedoxETL
{
    <#
    10=Successful
    20=Warnings
    30=Errors
    40=Failed
    50=Stopped
    60=Aborted
    0 = No error occurred (other commands than Execution)
    -1=Jedox Integrator was not reachable
    -2=Error on Client or Server side (e.g. Project/Job not found, Project not valid)
    #>
    param(
          [parameter(position=0)]
          $project,
          [parameter(position=1)]
          $job
         )
    $client = "I:\\Jedox\\Jedox Suite\\tomcat\\client"
    Set-Location -Path $client
    
    $rtvinfo = @(10,(Get-Date).ToString("yyyyMMdd_HHmmss"))
    $runETL = (Start-Process -FilePath etlclient.bat -ArgumentList "-sp PROFILES","-p $project","-j $job" -Wait -PassThru)
    $rtvinfo[0] = $runETL.ExitCode
    $rtvinfo[1] = $runETL.ExitTime.ToString("yyyyMMdd_HHmmss")
    $runETL.Close()

    return $rtvinfo
}

<#
HULFT送信バッチ呼び出す用
#>
function execHULFT
{
    param(
          [parameter(position=0)]
          $hulftid
         )
    $hulft = "I:\\hulft8\\binnt"
    Set-Location -Path $hulft
    $rtvinfo=@(0, "OK")

    $hulftsend = (Start-Process -FilePath utlsend.exe -ArgumentList "-f $hulftid","-sync" -Wait -PassThru)
    if($hulftsend.ExitCode -ne 0)
    {
        $rtvinfo[0]=$hulftsend.ExitCode
        $rtvinfo[1]=[String]::Format("HULFT処理{0}が失敗しました。終了コード【{1}】です",$h, $hulftsend.ExitCode)
    }
    $hulftsend.Close()

    return $rtvinfo

}

function TestF
{
    param(
          [parameter(position=0)]
          $Msg
         )
    
    Write-Host ("メッセージ：")
    Write-Host ($Msg)
}


