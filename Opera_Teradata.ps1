Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$VerbosePreference = "Continue"
$DebugPreference = "Continue"

#teradataに接続
[string] $dataSource = 'DataSource'  
$authentication = ("User Id={0};Password={1};" -f "USERNAME", "PASSWORD")
$factory = [System.Data.Common.DbProviderFactories]::GetFactory("Teradata.Client.Provider")
$connection = $factory.CreateConnection() 
$connection.ConnectionString = "Data Source = $dataSource;Connection Pooling Timeout=300;$authentication" 
$connection.Open()

#事前データを削除
$sqlCommand = "DELETE FROM TABLE1 WHERE calendar_ym = '201708' AND company_cd='XXXX'"
$command = $connection.CreateCommand()
$command.CommandText = $sqlCommand
$adapter = $factory.CreateDataAdapter()
$adapter.SelectCommand = $command
$dataset = new-object System.Data.DataSet
[void] $adapter.Fill($dataset)
$dataset.Tables | Select-Object -Expand Rows


# ライブラリ読み込み
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Data")
# H2DB接続
$H2DBconnectionString = "DSN=H2DB;uid=sa;pwd=sa;"
$odbcCon = New-Object System.Data.Odbc.OdbcConnection($H2DBconnectionString)
$odbcCon.Open()

# コマンドオブジェクト作成
$odbcCmd = New-Object System.Data.Odbc.OdbcCommand
$odbcCmd.Connection = $odbcCon

# コマンド実行（SELECT）
$odbcCmd.CommandText = "SELECT * FROM TABLE WHERE calendar_ym = '201708' AND company_cd='XXXX'"
$odbcReader = $odbcCmd.ExecuteReader()
while ($odbcReader.Read()) {
    $ITEM1 =$odbcReader["calendar_ym"].ToString()
    $ITEM2 =$odbcReader["company_cd"].ToString()
    $ITEM3 =$odbcReader["srcsysitem_cd"].ToString()
    $ITEM4 =$odbcReader["invcost_am"].ToString()

    $sqlCommand = "INSERT INTO TABLE "
    $sqlCommand += "(calendar_ym,company_cd,srcsysplant_cd,srcsysitem_cd,invcost_am,dataaddusr_id,dataaddprg_id,dataadd_yd,dataupdusr_id,dataupdprg_id,dataupd_yd,dataupd_rv) VALUES "
    $sqlCommand += "('$ITEM1', '$ITEM2', null, '$ITEM3', '$ITEM4', 'SHINGOU_RI', 'manual', CURRENT_TIMESTAMP(0), null, null, null, '0')"
    $command = $connection.CreateCommand()
    $command.CommandText = $sqlCommand
    $adapter = $factory.CreateDataAdapter()
    $adapter.SelectCommand = $command
    $dataset = new-object System.Data.DataSet
    [void] $adapter.Fill($dataset)
    $dataset.Tables | Select-Object -Expand Rows
}
$odbcReader.Dispose()

# コマンドオブジェクト破棄
$odbcCmd.Dispose()
$command.Dispose()
# DB切断
$odbcCon.Close()
$odbcCon.Dispose()
$connection.Dispose()
$connection.Close()





































