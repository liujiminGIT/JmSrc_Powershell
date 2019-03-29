# A1形式からR1C1形式に変換する関数
function Convert-A1ToR1C1($address){
  # 変換テーブル作成
  $alph = @{
    A=1;B=2;C=3;D=4;E=5;F=6;G=7;H=8;I=9;J=10;
    K=11;L=12;M=13;N=14;O=15;P=16;Q=17;R=18;S=19;T=20;
    U=21;V=22;W=23;X=24;Y=25;Z=26
  }

  # 列名を取得して1文字ずつの配列化
  $ary = ($address -split "\d")[0].ToCharArray()
  # 計算用に配列を反転
  $ary = $ary[($ary.length-1)..0]
  # 計算用一時変数
  $pow = 1
  $col = 0
  # 列名を数字に変換（26進数を10進数に変換）
  $ary | %{
    $col += ($alph[$_.ToString()] * $pow)
    $pow *= 26
  }

  # リターン用オブジェクト作成
  $ret = [PSCustomObject]@{
    row = ($address -split "\D")[-1]
    col = $col
  }
  return $ret
}

# R1C1形式からA1形式に変換する関数
function Convert-R1C1ToA1($row, $col){
  # 変換テーブル作成
  $alph = @(
    "Z","A","B","C","D","E","F","G","H","I","J",
    "K","L","M","N","O","P","Q","R","S","T",
    "U","V","W","X","Y"
  )
  # 計算用一時変数
  $ary = @()
  $div = $col
  # 列名を数字に変換
  for(;;){
    $mod = $div%26
    $div = [Math]::Floor($div/26)

    # 例外的に、余りが0の場合は商を1減らす
    if($mod -eq 0){
      $div -= 1
    }

    $ary += $alph[$mod]

    if($div -le 0){
      break
    }
  }
  # 逆順で求まるので配列を反転
  $ary = $ary[($ary.length-1)..0]

  # リターン用オブジェクト作成
  $ret = [PSCustomObject]@{
    row = $row
    col = $ary -join ""
  }
  return $ret
}


# 実行テスト
$obj_ary = @()
@(1..100) | %{
  $a1 = Convert-R1C1ToA1 1 $_
  $r1c1 = Convert-A1ToR1C1 ($a1.col+$a1.row)
  $obj = [PSCustomObject]@{
    R1C1 = ($r1c1.col)
    A1 = ($a1.col)
  }
  $obj_ary += $obj
}
@(650..710) | %{
  $a1 = Convert-R1C1ToA1 1 $_
  $r1c1 = Convert-A1ToR1C1 ($a1.col+$a1.row)
  $obj = [PSCustomObject]@{
    R1C1 = ($r1c1.col)
    A1 = ($a1.col)
  }
  $obj_ary += $obj
}
$obj_ary | ft -AutoSize