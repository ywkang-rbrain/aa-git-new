' ##############  필터하기  ##############


' WScript.Arguments.Item(0)             '파일명
' WScript.Arguments.Item(1)			  '시트명 
' WScript.Arguments.Item(2)		      '필터할 열번호 (A=1, B=2, C=3, D=4 ..........) 
' WScript.Arguments.Item(3)			  '필터할 단어     /   구분자(WScript.Arguments.Item(4))를 읽어서 구분하여 받음  /  Ex) A1,A3,A5 == 3개 단어 필터
' WScript.Arguments.Item(4)			  '필터구분자 : 사용자가 지정     /    1개열에 필터할 단어가 2개 이상일때 구분하는 구분자  Ex) 가,다,마

Dim Input_Tye
Dim ColumnNum
Dim wb
Dim wks

'숫자인지 문자인지 구분하기 ------------------------------------------------------------------------
If IsNumeric(WScript.Arguments.Item(2)) Then
    Input_Tye = "Numeric"
	ColumnNum = Cint(WScript.Arguments.Item(2))
Else
    Input_Tye = "String"

End If

'문자인 경우 숫자로 변환하기 ------------------------------------------------------------------------
If Input_Tye = "String" Then

			Dim i
			Dim len_col
			Dim num_temp
			Dim str_digit 

			col_name = UCase(WScript.Arguments.Item(2))    
			For i = 1 To Len(col_name)
				str_digit = Mid(col_name, i, 1)
				num_temp = Asc(str_digit) - 64
				NumFromXLColumn = NumFromXLColumn + num_temp * 26 ^ (Len(col_name) - i)
			Next 		   

			ColumnNum = Cint(NumFromXLColumn)

End If


'필터하기 ------------------------------------------------------------------------------------------------
Set wb = GetObject(WScript.Arguments.Item(0))
Set wks = wb.Worksheets(WScript.Arguments.Item(1))
ArgCnt = WScript.arguments.count  '한개 더 카운트 됨 


if ArgCnt = "5" then '한 단어만 필터할 경우
	wks.Range("A1").AutoFilter ColumnNum, Split(WScript.Arguments.Item(3), WScript.Arguments.Item(4)), 1  '1번은 <> 포함하지않음  /  * 포함 필터 가능

elseif ArgCnt = "6" then ' 여러 단어로 필터할 경우
	wks.Range("A1").AutoFilter ColumnNum, Split(WScript.Arguments.Item(3), WScript.Arguments.Item(4)), 7  '7번은 멀티 필터 가능 / 한단어 필터도 가능

else '옵션이 잘못 설정된 경우 - 에러창 10초 동안만
	Dim WSH
	Set WSH = CreateObject("WScript.Shell") 
	WSH.Popup "옵션 개수가 잘못되었습니다."  & Chr(10) & Chr(13) & "AA에서 필터 옵션을 다시 한번 확인하세요."  & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "10 초 후 자동으로 닫힙니다", 10, "옵션 개수 오류", vbInformation 
	Set WSH = Nothing

end if


Set wks = Nothing
Set wb = Nothing 

Err.Clear 


'xlAnd                         1           And
'xlOr                           2           Or
'xlBottom10Items         4           하위 10개
'xlBottom10Percent      6           하위 10%
'xlFilterCellColor          8            셀 색상
'xlFilterDynamic          11            다이나믹
'xlFilterFontColor        9             폰트 색
'xlFilterIcon                10            아이콘
'xlFilterValues             7             값
'xlTop10Items             3             상위 10개
'xlTop10Percent         5             상위 10%