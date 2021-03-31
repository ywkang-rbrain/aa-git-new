'#####  복사 영역 지정해서 복사하기 #####

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)  
'시트명 = WScript.Arguments.Item(1)
'복사범위 = WScript.Arguments.Item(2)


	On Error Resume Next


Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False
  
wks.Range(WScript.Arguments.Item(2)).Copy

Set wks = nothing
Set wb = nothing 

Err.Clear