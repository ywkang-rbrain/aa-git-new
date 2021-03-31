'#####  텍스트 숫자 변환 #####

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)
'시트명 = WScript.Arguments.Item(1)
'영문열번호 = WScript.Arguments.Item(2)

'한번에 한 열씩만 변경 가능 ---> 2개열 이상은 각각 한번씩 변경하기

On Error Resume Next

Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

'문자인 경우 숫자로 변환하기 ------------------------------------------------------------------------
col_name = UCase(WScript.Arguments.Item(2))    
For i = 1 To Len(col_name)
	str_digit = Mid(col_name, i, 1)
	num_temp = Asc(str_digit) - 64
	NumFromXLColumn = NumFromXLColumn + num_temp * 26 ^ (Len(col_name) - i)
Next 		   
ColumnNum = Cint(NumFromXLColumn)
'문자인 경우 숫자로 변환하기 ------------------------------------------------------------------------

wks.Cells(ColumnNum).EntireColumn.TextToColumns

Set wks = nothing
Set wb = nothing


Err.Clear      '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제
