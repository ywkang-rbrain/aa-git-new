'##### 주문 수량 수정  #####

'WScript.Arguments.Item(0)					기준단위
'WScript.Arguments.Item(1)					발주수량
'WScript.Arguments.Item(2)					반올림할 %   -->   ex) 20, 30 반드시 숫자만 입력 

'ex)
'주문가능단위 대비 발주수량 30% 미만 (1~29%까지)인 경우 0, 이상인 경우 주문가능단위로 변경
'주문단위 40, (발주수량 10 -> 0 / 발주수량 20 -> 40)


'remainder
On Error Resume Next

Dim vBaseUnit					'기준단위
Dim vOrderQuantity			'발주수량
Dim vShare 
Dim vDivision 
Dim vRemainder 


vBaseUnit = Cint(WScript.Arguments.Item(0))
vOrderQuantity = Cint(WScript.Arguments.Item(1))

vDivision = (vOrderQuantity / vBaseUnit)             '나눈값
vShare = Fix(vOrderQuantity / vBaseUnit)            '몫
vRemainder = vBaseUnit Mod vOrderQuantity      '나머지

WScript.StdOut.Write (vRemainder)  'vResultValue Return

Err.Clear      '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제