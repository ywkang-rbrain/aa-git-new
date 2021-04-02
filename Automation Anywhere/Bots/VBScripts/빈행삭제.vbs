'#####빈 행 삭제 #####

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)
'시트명 = WScript.Arguments.Item(1)
'검색할 열 = WScript.Arguments.Item(2)		검색할 열  숫자로 Ex) A=1, B=2, C=3................ 
'시작행 = WScript.Arguments.Item(3)	    	숫자만 입력
'마지막행 = WScript.Arguments.Item(4)		숫자만 입력

'Ex) "C:\aaa.xlsx" "Sheet1" "A" "2" "200"

On Error Resume Next

Dim wb
Dim wks
Dim i
Dim j
Dim ColumnNum
Dim StartRow
Dim EndRow 
Dim vColumn

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))

'문자인 경우 숫자로 변환하기 ------------------------------------------------------------------------
col_name = UCase(WScript.Arguments.Item(2))    
For i = 1 To Len(col_name)
	str_digit = Mid(col_name, i, 1)
	num_temp = Asc(str_digit) - 64
	NumFromXLColumn = NumFromXLColumn + num_temp * 26 ^ (Len(col_name) - i)
Next 		   
ColumnNum = Cint(NumFromXLColumn)
'문자인 경우 숫자로 변환하기 ------------------------------------------------------------------------

StartRow = Cint(WScript.Arguments.Item(3))    
EndRow = Cint(WScript.Arguments.Item(4))       

j=1
For i = StartRow to EndRow          
  
  If Trim(wks.Cells(i, ColumnNum).Value) = "" And j <> EndRow Then            
		wks.Rows(i).Delete
		i = i-1         
  ElseIf  EndRow = j Then
		If Trim(wks.Cells(i, ColumnNum).Value) = "" Then
			wks.Rows(i).Delete
		End If
		Exit For	
  End If
  j = j + 1

Next 

Err.Clear      '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제


