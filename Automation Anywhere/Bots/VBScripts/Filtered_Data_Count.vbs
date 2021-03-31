'#####  필터된 행수 구하기 / 행갯수 리턴 ##### 
' 변수로 받을 경우 에러가 있어서 소스에 직접 Parameter 입력
'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력


'파일전체경로 = WScript.Arguments.Item(0)  
'시트명 = WScript.Arguments.Item(1)

'######################################################################################
'반드시 데이터가 들어있는 한열만 지정   /   ex)A2:J50(X) A~J열의 모든 행들을 합쳐서 계산하여 행수 안맞음
'######################################################################################
'전체 범위  = WScript.Arguments.Item(2)  ex) A2:A10  /  B2:B50   /  C2:C60  

On Error Resume Next


Dim wb
Dim wks
Dim Filtered_Data_Count

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

Filtered_Data_Count = wks.Range(WScript.Arguments.Item(2)).SpecialCells(12).Count

If Filtered_Data_Count = "" then
	Filtered_Data_Count = 0
end if

WScript.StdOut.Write (Filtered_Data_Count)  '찾은 Filtered_Data_Count값 Return

Set wks = nothing
Set wb = nothing


Err.Clear  