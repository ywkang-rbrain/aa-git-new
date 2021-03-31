'#####  필터 해제 하기 #####

' 변수로 받을 경우 에러가 있어서 소스에 직접 Parameter 입력
'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

' WScript.Arguments.Item(0)             '파일명
' WScript.Arguments.Item(1)			  '시트명 

On Error Resume Next


Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

wks.AutoFilterMode = False

Set wks = Nothing
Set wb = Nothing 


Err.Clear