'#####  필터된 모든행 삭제하기 ##### 
'1행 헤더부분을 삭제를 안할려면  WScript.Arguments.Item(2)의 값을 2행부터 지정


' 변수로 받을 경우 에러가 있어서 소스에 직접 Parameter 입력
'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력


'파일전체경로 = WScript.Arguments.Item(0)  
'시트명 = WScript.Arguments.Item(1)
'삭제할 범위  = WScript.Arguments.Item(2)   ex) 2:10
On Error Resume Next

Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

wks.Rows(WScript.Arguments.Item(2)).SpecialCells(12).EntireRow.Delete

Set wks = nothing
Set wb = nothing


Err.Clear