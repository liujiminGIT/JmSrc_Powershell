<#
配置手順
■STEP1
　TimeRecord_Start.ps1を以下場所に配置する
　C:\Windows\System32\GroupPolicy\Machine\Scripts\Startup
  TimeRecord_End.ps1を以下場所に配置する
　C:\Windows\System32\GroupPolicy\Machine\Scripts\Shutdown
■STEP2
　Cドライバ直下に「Work_Record」のフォルダを作成する
■STEP3
　「WIN+R」キーを押して、「ファイル名を指定して実行」のところ、「gpedit.msc」を入力して、「Enter」を押す
■STEP4
　「ローカルグループポリシーエディター」Windowが表示される。
　ローカルグループポリシー
　　→コンピューターの構成
　　　→Windowsの設定
　　　　→スクリプト（スタートアップ/シャットダウン）
　　を選択する。
■STEP5
　「スタートアップ」をダブルクリックして、
　「スクリプト」タブの「追加」ボタンを押して、
　参照を押して、
　「TimeRecord_Start.ps1」を選択する
　OK保存
■STEP6
　「シャットダウン」をダブルクリックして、
　「スクリプト」タブの「追加」ボタンを押して、
　参照を押して、
　「TimeRecord_End.ps1」を選択する
　OK保存
-------------------以上--------------------------

パソコンシャットダウン／起動する度に、C:\Work_Recordしたの
「yyyymm_Recod.log」のファイルに、タイムスタンプが記録される
ファイルが無かったら自動作成される。
月単位でファイルを持つ。

例：
---------------------------------------------
SATRT TIME	END TIME
2016/05/02 09:23:41	2016/05/02 20:04:45
2016/05/06 09:19:45	2016/05/06 20:23:47
----------------------------------------------

#>

<#
TimeRecord_Start.ps1
#>

$fileName = "C:\Work_Record\"+(Get-Date).ToString("yyyyMM")+"_Recod_.log"
$head = "SATRT TIME`tEND TIME"
#$dt_start = "`r`n"+(Get-Date).AddMinutes(-8).ToString("yyyy/MM/dd HH:mm:ss")+"`t"
$dt_start = "`r`n"+[Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem).LastBootUpTime).ToString("yyyy/MM/dd HH:mm:ss")+"`t"
#$dt_end = (Get-Date).ToString("yyyy/MM/dd HH:mm:ss")
if(!(Test-Path $fileName))
{
    New-Item -Path $fileName -Value $head
}

#Start
$dt_start | Add-Content -Path $fileName -NoNewline

#End
#$dt_end | Add-Content -Path $fileName -NoNewline
<#########################################################################>

<#
TimeRecord_End.ps1
#>
$fileName = "C:\jm_work\8110_WorkRecord\"+(Get-Date).ToString("yyyyMM")+"_Recod_.log"
$head = "SATRT TIME`tEND TIME"
#$dt_start = "`r`n"+(Get-Date).ToString("yyyy/MM/dd HH:mm:ss")+"`t"
$dt_end = (Get-Date).AddMinutes(8).ToString("yyyy/MM/dd HH:mm:ss")
if(!(Test-Path $fileName))
{
    New-Item -Path $fileName -Value $head
}

#Start
#$dt_start | Add-Content -Path $fileName -NoNewline

#End
$dt_end | Add-Content -Path $fileName -NoNewline

<#########################################################################>


<#
Windowsのログから、パソコンの起動終了時間一括抽出する
#>


$q='
<QueryList>
  <Query Id="0" Path="System">
    <Select Path="System">*[System[(EventID=6005 or EventID=6006)]]</Select>
  </Query>
</QueryList>'
$events = Get-WinEvent -FilterXml $q
$i=-1
$outfile = "C:\jm_work\PC_StartStopTime_"+((Get-Date).ToString("yyyyMMdd_HHmmss"))+".txt"
Write-Output ("6005_start`t6006_end`twork(hour)") >> $outfile

while ( $i+1 -lt $events.length ) {
  if($i -eq -1)
  {
    $StartTime = $events[0].TimeCreated
    $StopTime = $null 
    $UpTime = [datetime]::Now - $events[0].TimeCreated
  }
  else{
    $StartTime = $events[$i+1].TimeCreated   #6006 停止時刻 イベント ログ サービスが停止されました。
    $StopTime = $events[$i].TimeCreated      #6005 開始時刻 イベント ログ サービスが開始されました。
    $UpTime = $events[$i].TimeCreated - $events[$i+1].TimeCreated
    $uptime_abc = $UpTime.TotalMinutes / 60
  }
  Write-Output ([String]::Format("{0:yyyy/MM/dd HH:mm:ss}`t{1:yyyy/MM/dd HH:mm:ss}`t{2:000.00}", $StartTime,$StopTime,$uptime_abc)) >> $outfile

  <#
  if($i -eq -1)
  {
    [PSCustomObject]@{
    StartTime = $events[0].TimeCreated;
    StopTime = $null ;
    UpTime = [datetime]::Now - $events[0].TimeCreated
    }
  }
  else{
  [PSCustomObject]@{
    StartTime = $events[$i+1].TimeCreated;   #6006 開始時刻 イベント ログ サービスが停止されました。
    StopTime = $events[$i].TimeCreated ;     #6005 開始時刻 イベント ログ サービスが開始されました。
    UpTime = $events[$i].TimeCreated - $events[$i+1].TimeCreated
    }
  }
  #>
  $i += 2
}


<#########################################################################>