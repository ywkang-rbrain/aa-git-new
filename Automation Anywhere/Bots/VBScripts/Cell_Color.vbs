'#####  복사된 영역 붙여넣기  /  필터된 영역 붙여넣기 도 가능 ##### 

'변수로 받을 경우 에러가 있어서 소스에 직접 Parameter 입력
'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)  
'시트명 = WScript.Arguments.Item(1)
'범위 = WScript.Arguments.Item(2)
'WScript.Arguments.Item(3)        번호로 입력  Ex) 1=검정, 2=흰색, 3=빨강, 4=초록, 5=파랑, 6=노랑 .........  56까지 있음 
'																		 정확한 색상은 인터넷에 "Interior.ColorIndex"로 검색하면 나옴


On Error Resume Next

Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

wks.Range(WScript.Arguments.Item(2)).Interior.ColorIndex = WScript.Arguments.Item(3)

Set wks = nothing
Set wb = nothing

Err.Clear