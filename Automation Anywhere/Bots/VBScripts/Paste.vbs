'#####  복사된 영역 붙여넣기 #####

'WScript.Arguments(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments(0)
'시트명 = WScript.Arguments(1)
'복사범위 = WScript.Arguments(2)
'v구분자 =   WScript.Arguments(3)
'vOption = WScript.Arguments(4)

On Error Resume Next

Dim wb
Dim wks
Dim vRange
Dim vOption    

Set wb = GetObject(WScript.Arguments(0))
Set wks = wb.Worksheets(WScript.Arguments(1))
wb.Application.DisplayAlerts = False
vRange = WScript.Arguments(2) 
vOption = WScript.Arguments(3)

'WScript.Sleep (1000)	'복사전에 1초 Sleep

	'같은파일 +  Ctrl+V		/   형식유지												  / 테두리 : O   /  색 : O
	IF vOption = "같은파일_전체" Then
		wks.Range(vRange).PasteSpecial 1

	'같은파일 + 값만붙여넣기		/ 	형식유지 					    			  / 테두리 : X   /  색 : X
	ElseIF vOption = "같은파일_값만" Then
		wks.Range(vRange).PasteSpecial 12

	'다른파일 +  Ctrl+V		/   형식유지												  / 테두리 : O   /  색 : O
	ElseIF vOption = "다른파일_전체" Then
		wks.Range(vRange).PasteSpecial 8	

	'다른파일 +  값만붙여넣기    /   형식유지  + 창뜸							  / 테두리 : O   /  색 : O    + (창뜸+ 창뜸+ 창뜸+ 창뜸+ 창뜸+ 창뜸)  -->  창은 다시 열지 않음으로 처리
	ElseIF vOption = "다른파일_값만" Then
		wks.Range(vRange).PasteSpecial -4163

	'조건 이외의 모든 경우    /   값만 복사 											  / 테두리 : X   /  색 : X
	Else
		wks.Range(vRange).PasteSpecial 12

	End If 	

Set wks = nothing
Set wb = nothing

Err.Clear