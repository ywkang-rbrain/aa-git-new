'==========  Find로 찾은 검색어의 Address 찾기(한개만 찾음)  ================

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'파일전체경로 = WScript.Arguments.Item(0)  
'시트명 = WScript.Arguments.Item(1)
'검색할 범위 = WScript.Arguments.Item(2)				' Ex) A:A(O)   /  A1:A500(O)  /  A1:Z100(O)    |  A(X)
'검색어 = WScript.Arguments.Item(3)						
'일치율 = WScript.Arguments.Item(4)						' 1 = 100%일치  /  2 = 부분일치
'찾는 방향 = WScript.Arguments.Item(5)					' 1= 위에서 아래  /  2 = 아래서 위
'Return 받을 형식 = WScript.Arguments.Item(6)			' 문열(문자 열 만받음) / 숫열(숫자로 열 번호만 받음)  / 숫행(숫자 행 번호만 받음) /  옵션을 입력하지 않으면(E6과 같이 기본 형태로 받음)   

'------------------------------------------------------------------------------------------------------------------------------------------------------------
'			1. 파일경로  2. 시트명	 3. 범위		4. 검색어		5. ( 1=완전일치, 2=부분일치)  6. 찾는방향(1=위에서아래, 2=아래서위)  7. Return 받을 형식
'옵션1  ex)	 "C:\aa.xlsx"	 "Sheet1"	 "A:A"  "사과"  "1"  "1" 
'옵션2  ex)	 "C:\aa.xlsx"	 "Sheet1"	 "A:A"  "사과"  "1"  "1" "숫행"
'------------------------------------------------------------------------------------------------------------------------------------------------------------

'==============================================
'검색된 경우 = Return 값 사용
'==============================================
'검색이 안된 경우 = "Not Found"를 Return함
'AA에서 
'If "Not Found" 하고      Else로 분기해서 사용
'==============================================

On Error Resume Next

Dim wb
Dim wks

Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
wb.Application.DisplayAlerts = False

Dim FoundCell
Set FoundCell = wks.Range(WScript.Arguments.Item(2)).Find (WScript.Arguments.Item(3),,,WScript.Arguments.Item(4),,WScript.Arguments.Item(5))

If FoundCell Is Nothing Then
   Final_Value = "Not Found"
   ColRow = Final_Value
Else
   Final_Value = FoundCell.Address
   Final_Value_Split = Split(Final_Value,"$")
   vColumn = Final_Value_Split(1)
   vRow = Final_Value_Split(2)
End If

If Final_Value = "Not Found" Then
	ColRow = "Not Found"
Else	'문열 / 숫열 / 숫행 / 기본(옵션생략)

		If WScript.Arguments.Item(6) = "문열" Then
			ColRow= vColumn

		ElseIf WScript.Arguments.Item(6) = "숫열" Then
				'문자열을 숫자로 변환 ---------------------------------------------------------------------------
				col_name = vColumn
				For i = 1 To Len(col_name)
					str_digit = Mid(col_name, i, 1)
					num_temp = Asc(str_digit) - 64
					NumFromXLColumn = NumFromXLColumn + num_temp * 26 ^ (Len(col_name) - i)
				Next 		   

				ColumnNum = Cint(NumFromXLColumn)
				'문자열을 숫자로 변환 ---------------------------------------------------------------------------
			ColRow= ColumnNum

		ElseIf WScript.Arguments.Item(6) = "숫행" Then
			ColRow= vRow

		Else
			ColRow = vColumn & vRow

		End if
	
End if

WScript.StdOut.Write (ColRow)  

Set wks = nothing
Set wb = nothing

Err.Clear       '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제