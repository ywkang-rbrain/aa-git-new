'#####  열 너비 조절  #####

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)
'시트명 = WScript.Arguments.Item(1)
'열 = WScript.Arguments.Item(2)						ex) A,  A:C 
'고정/자동 = WScript.Arguments.Item(3)
'고정일경우 = WScript.Arguments.Item(4)				' 열너비 값

On Error Resume Next

Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

If WScript.Arguments.Item(3) = "고정" Then
	wks.Columns(WScript.Arguments.Item(2)).ColumnWidth = WScript.Arguments.Item(4)
ElseIf WScript.Arguments.Item(3) = "자동" Then
	wks.Columns(WScript.Arguments.Item(2)).Autofit
End if


Set wks = nothing
Set wb = nothing

Err.Clear       '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제