''#####  선택 영역 자동 채우기 #####

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)  
'시트명 = WScript.Arguments.Item(1)
'붙여넣을 기준위치 = WScript.Arguments.Item(2)
'붙여넣을 범위 = WScript.Arguments.Item(3)

On Error Resume Next



Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

wks.Range(WScript.Arguments.Item(2)).Autofill wks.Range(WScript.Arguments.Item(3))

Set wks = nothing
Set wb = nothing

Err.Clear