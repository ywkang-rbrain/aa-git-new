'#####  표시 형식 변경 #####

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)
'시트명 = WScript.Arguments.Item(1)
'범위 = WScript.Arguments.Item(2)
'포맷형식 = WScript.Arguments.Item(3)		일반 / 텍스트 / 숫자 / 회계 / 콤마(천단위) / 퍼센트 / 강조색1 / 사용자지정
'사용자지정인 경우 = WScript.Arguments.Item(4)		'사용자지정인 경우 /  사용자 지정 숫자  Ex) "#,###.00;[Red](#,###.00);0.00"


On Error Resume Next

Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False			'알림창 X

If WScript.Arguments.Item(3) = "일반" Then 
'	wks.Range(WScript.Arguments.Item(2)).Style = "Normal"					'일반 입력 - 방법1	
	wks.Range(WScript.Arguments.Item(2)).NumberFormat = "General"		'일반 입력 - 방법2		
ElseIf WScript.Arguments.Item(3) = "텍스트" Then 
	wks.Range(WScript.Arguments.Item(2)).NumberFormatLocal = "@"			'텍스트 입력
ElseIf WScript.Arguments.Item(3) = "숫자" Then 
	wks.Range(WScript.Arguments.Item(2)).NumberFormat =  "0"					'숫자
ElseIf WScript.Arguments.Item(3) = "회계" Then 
	wks.Range(WScript.Arguments.Item(2)).Style = "Currency [0]"						'통화
ElseIf WScript.Arguments.Item(3) = "콤마" Then 
	wks.Range(WScript.Arguments.Item(2)).Style = "Comma [0]"					'콤마(천단위 콤마 찍기)
ElseIf WScript.Arguments.Item(3) = "퍼센트" Then 
	wks.Range(WScript.Arguments.Item(2)).Style = "Percent"						'퍼센트
ElseIf WScript.Arguments.Item(3) = "강조색1" Then 
	wks.Range(WScript.Arguments.Item(2)).Style = "강조색1"						'강조색1
ElseIf WScript.Arguments.Item(3) = "사용자지정" Then 
	wks.Range(WScript.Arguments.Item(2)).NumberFormat =  WScript.Arguments.Item(4)	'사용자 지정 숫자  Ex) "#,###.00;[Red](#,###.00);0.00"
End If


Set wks = nothing
Set wb = nothing


Err.Clear      '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제

