# 공통: 숫자 읽기, 거래처 이름 짝(내 이름 → 그쪽 거래처)
function N([string]$s) { $x = $s -replace '[^\d\.\-]', ''; if ($x -match '^-?\d+(\.\d+)?$') { [decimal]$x } else { [decimal]0 } }
$ALIAS = @{
  '귀뚜라미범양냉방' = '귀뚜라미범양냉방'; '귀뚜라미범양' = '귀뚜라미범양냉방'; '부경테크윈' = '부경테크원'; 'ATLASKOREA(국·해외)' = '아트라스콥코'; 'DN솔루션즈' = '디엔솔루션즈'
  'STX엔진제1사업장' = 'STX엔진'; 'STX엔진민수' = 'STX엔진'; 'STX엔진방산' = 'STX엔진'; '명화' = '명화공업'; '엘에스엠트론' = 'LS엠트론'; '현대로템방산' = '현대로템'; '현대로템철차' = '현대로템'; '현대트랜시스중기' = '현대트랜시스'
  'ALLISON' = 'ALLISON'; 'BERGENENG' = 'BERGEN ENGINE'; 'BERGENENGINE' = 'BERGEN ENGINE'; 'JSTECH(코마츠)' = '코마츠(JS TECH)'; '코마츠(JSTECH)' = '코마츠(JS TECH)'; 'MERCURYC' = 'MERCURY'; 'MERCURYCHINA' = 'MERCURY'; 'MERCURYU' = 'MERCURY'; 'MERCURY' = 'MERCURY'
  'SHINKO' = '신코'; 'SINKOINDLTD' = '신코'; 'EMERSON' = 'EMERSON'; 'POLARIS' = 'POLARIS'; 'DOMETIC' = 'DOMETIC'; 'AMAT' = 'AMAT'; 'SAM' = 'SAM'; 'MEPPI' = 'MEPPI'
}
function Key([string]$nm) { $k = ($nm -replace '㈜|\(주\)|주식회사|\s', '').ToUpper(); if ($ALIAS.ContainsKey($k)) { return $ALIAS[$k] }; return $k }
