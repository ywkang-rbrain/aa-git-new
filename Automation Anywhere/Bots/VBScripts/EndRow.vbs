'#####  끝행값 찾아서 값 리턴하기 ##### 

' 변수로 받을 경우 에러가 있어서 소스에 직접 Parameter 입력
'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력


'파일전체경로 = WScript.Arguments.Item(0)  
'시트명 = WScript.Arguments.Item(1)
'끝행을 찾을 열  = WScript.Arguments.Item(2)      ex) A, B, C, D ......  

On Error Resume Next

Dim wb
Dim wks
Dim EndRow

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

EndRow = wks.Range(WScript.Arguments.Item(2) & wks.Rows.Count).End(-4162).Row   'xlUp = -4162   /   xlDown = -4121

If EndRow = 1 Then
	If Trim(wks.Range(WScript.Arguments.Item(2) & 1).Value) = "" Then
		EndRow = 0
	Else
		EndRow = 1
	End if
End if

WScript.StdOut.Write (EndRow)  '찾은 EndRow값 Return

Set wks = nothing
Set wb = nothing

Err.Clear       '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제