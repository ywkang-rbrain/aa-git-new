'#####  Create_Workbook  #####

'WScript.Arguments.Item(0)와 같이 AA에서 넘기는 변수를 vFileName과 같은 변수로 지정해서 할경우 에러가 있어서 각 위치에 직접 입력

'WScript.Arguments.Item(0)					파일전체경로
'WScript.Arguments.Item(1)					확장자

On Error Resume Next

Dim App
Dim wb
Dim wks
Dim FileFormat

Set App = CreateObject("Excel.Application") 
App.Application.DisplayAlerts = False			'알림창 X

Set wb = App.Workbooks.Add

If WScript.Arguments.Item(1) = "xlsb" Then
	FileFormat = "50"
ElseIf WScript.Arguments.Item(1) = "xlsx" Then
	FileFormat = "51"
ElseIf WScript.Arguments.Item(1) = "xlsm" Then
	FileFormat = "52"
ElseIf WScript.Arguments.Item(1) = "xls" Then
	FileFormat = "56"
End If

wb.SaveAs WScript.Arguments.Item(0) , FileFormat
wb.Close

Set App = nothing
Set wb = nothing

Err.Clear      '에러창 안뜨게 하는 명령어  :   개발 단계에서는 주석 처리  /  운영에서는 반드시 주석 해제


'50 = xlExcel12 (Excel Binary Workbook in 2007-2016 with or without macro's, xlsb)
'51 = xlOpenXMLWorkbook (without macro's in 2007-2016, xlsx)
'52 = xlOpenXMLWorkbookMacroEnabled (with or without macro's in 2007-2016, xlsm)
'56 = xlExcel8 (97-2003 format in Excel 2007-2016, xls)