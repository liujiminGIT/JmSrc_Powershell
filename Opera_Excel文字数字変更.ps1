# A1�`������R1C1�`���ɕϊ�����֐�
function Convert-A1ToR1C1($address){
  # �ϊ��e�[�u���쐬
  $alph = @{
    A=1;B=2;C=3;D=4;E=5;F=6;G=7;H=8;I=9;J=10;
    K=11;L=12;M=13;N=14;O=15;P=16;Q=17;R=18;S=19;T=20;
    U=21;V=22;W=23;X=24;Y=25;Z=26
  }

  # �񖼂��擾����1�������̔z��
  $ary = ($address -split "\d")[0].ToCharArray()
  # �v�Z�p�ɔz��𔽓]
  $ary = $ary[($ary.length-1)..0]
  # �v�Z�p�ꎞ�ϐ�
  $pow = 1
  $col = 0
  # �񖼂𐔎��ɕϊ��i26�i����10�i���ɕϊ��j
  $ary | %{
    $col += ($alph[$_.ToString()] * $pow)
    $pow *= 26
  }

  # ���^�[���p�I�u�W�F�N�g�쐬
  $ret = [PSCustomObject]@{
    row = ($address -split "\D")[-1]
    col = $col
  }
  return $ret
}

# R1C1�`������A1�`���ɕϊ�����֐�
function Convert-R1C1ToA1($row, $col){
  # �ϊ��e�[�u���쐬
  $alph = @(
    "Z","A","B","C","D","E","F","G","H","I","J",
    "K","L","M","N","O","P","Q","R","S","T",
    "U","V","W","X","Y"
  )
  # �v�Z�p�ꎞ�ϐ�
  $ary = @()
  $div = $col
  # �񖼂𐔎��ɕϊ�
  for(;;){
    $mod = $div%26
    $div = [Math]::Floor($div/26)

    # ��O�I�ɁA�]�肪0�̏ꍇ�͏���1���炷
    if($mod -eq 0){
      $div -= 1
    }

    $ary += $alph[$mod]

    if($div -le 0){
      break
    }
  }
  # �t���ŋ��܂�̂Ŕz��𔽓]
  $ary = $ary[($ary.length-1)..0]

  # ���^�[���p�I�u�W�F�N�g�쐬
  $ret = [PSCustomObject]@{
    row = $row
    col = $ary -join ""
  }
  return $ret
}


# ���s�e�X�g
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