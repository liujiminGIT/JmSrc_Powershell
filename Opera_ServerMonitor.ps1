<#
インストール済みのアプリケーション一覧
#>
Get-WmiObject Win32_Product | Select-Object Name,Vendor,Version,Caption | Where-Object {$_.name -Match ".*jedox.*"}


<#########################################################################>

<#
パフォーマンス状況監視
#>
#Get-Counter -ListSet Memory | Select-Object -ExpandProperty Paths

$counters = "\Memory\Committed Bytes","\Memory\Pages/sec","\Network Interface(*)\Bytes Total/sec","\Paging File(*)\% Usage","\PhysicalDisk(*)\% Disk Time","\Processor(*)\% Processor Time","\System\Processor Queue Length","\Memory\Available MBytes"

Get-Counter -Counter $counters -Continuous | % {
    $p = New-Object PSObject | Add-Member -PassThru NoteProperty TimeStamp $_.TimeStamp
    $_.CounterSamples | % { $p | Add-Member NoteProperty $_.Path $_.CookedValue }
    $p
}
#Get-Alias

<#########################################################################>


<#
OS情報取得
#>

$ReturnData = New-Object PSObject | Select-Object HostName,Manufacturer,Model,SN,CPUName,PhysicalCores,Sockets,MemorySize,DiskInfos,OS

$Win32_BIOS = Get-WmiObject Win32_BIOS
$Win32_Processor = Get-WmiObject Win32_Processor
$Win32_ComputerSystem = Get-WmiObject Win32_ComputerSystem
$Win32_OperatingSystem = Get-WmiObject Win32_OperatingSystem

# ホスト名
$ReturnData.HostName = hostname

# メーカー名
$ReturnData.Manufacturer = $Win32_BIOS.Manufacturer

# モデル名
$ReturnData.Model = $Win32_ComputerSystem.Model

# シリアル番号
$ReturnData.SN = $Win32_BIOS.SerialNumber

# CPU 名
$ReturnData.CPUName = @($Win32_Processor.Name)[0]

# 物理コア数
$PhysicalCores = 0
$Win32_Processor.NumberOfCores | % { $PhysicalCores += $_}
$ReturnData.PhysicalCores = $PhysicalCores

# ソケット数
$ReturnData.Sockets = $Win32_ComputerSystem.NumberOfProcessors

# メモリーサイズ(GB)
$Total = 0
Get-WmiObject -Class Win32_PhysicalMemory | % {$Total += $_.Capacity}
$ReturnData.MemorySize = [int]($Total/1GB)

# ディスク情報
[array]$DiskDrives = Get-WmiObject Win32_DiskDrive | ? {$_.Caption -notmatch "Msft"} | sort Index
$DiskInfos = @()
foreach( $DiskDrive in $DiskDrives ){
    $DiskInfo = New-Object PSObject | Select-Object Index, DiskSize
    $DiskInfo.Index = $DiskDrive.Index              # ディスク番号
    $DiskInfo.DiskSize = [int]($DiskDrive.Size/1GB) # ディスクサイズ(GB)
    $DiskInfos += $DiskInfo
}
$ReturnData.DiskInfos = $DiskInfos

# OS 
$OS = $Win32_OperatingSystem.Caption
$SP = $Win32_OperatingSystem.ServicePackMajorVersion
if( $SP -ne 0 ){ $OS += "SP" + $SP }
$ReturnData.OS = $OS

return $ReturnData



<#########################################################################>