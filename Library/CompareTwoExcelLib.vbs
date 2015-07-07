Option Explicit

''' #########################################################
''' <summary>
''' A library to Compare two excel sheets cell by cell and Highlight the different by red color
''' </summary>
''' <remarks></remarks>
''' <example>
''' CompareTwoExcelLib.Compare "C:\Config.xls", "C:\Config2.xls"
''' CompareTwoExcelLib.CompareIncludingHeaderRow "C:\Config.xls", "C:\Config2.xls"
''' CompareTwoExcelLib.CompareSpecifiedColumns "C:\Config.xls", "C:\Config2.xls", Array(1,2,4,8)
''' CompareTwoExcelLib.CompareSpecifiedColumnsIncludingHeaderRow "C:\Config.xls", "C:\Config2.xls", Array(1,2,3)
''' </example>
''' #########################################################

Class ClsCompareTwoExcelLib

	''' <summary>
    ''' Region Excel.Application instance created in Class_Initialize
    ''' </summary>
    ''' <remarks></remarks>
	Private xlsApp
	
	''' <summary>
    ''' Class Initialization procedure. Creates Excel Singleton.
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
		
		Dim bCreated : bCreated = False
		
		If IsObject(xlsApp) Then
			If Not xlsApp Is Nothing Then
				If TypeName(xlsApp) = "Application" Then
					bCreated = True
				End If
			End If
		End If
		
		If Not bCreated Then 
			On Error Resume Next
				Set xlsApp = GetObject("", "Excel.Application")

				If Err.Number <> 0 Then
					Err.Clear

					Set xlsApp = CreateObject("Excel.Application")
				End If
				
				If Err.Number <> 0 Then
					MsgBox "Please install Excel before using ExcelLib", vbOKOnly, "Excel.Application Exception!"
					Err.Clear
					Exit Sub
				End If
			On Error Goto 0
		End If
		
	End Sub
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
	
	 	xlsApp.Quit
		Set xlsApp = Nothing
		
	End Sub

	''' <summary>
    ''' Compare 2 Excel sheets cell by cell including Header Row
    ''' </summary>
    ''' <param name="Workbook1" type="string">Location to the Excel WorkBook 1</param>
    ''' <param name="Workbook2" type="string">Location to the Excel WorkBook 2</param>
    ''' <returns>True/False</returns>
	Public Function CompareIncludingHeaderRow(Byval Workbook1, Byval Workbook2)

		CompareIncludingHeaderRow = CompareExcel(Workbook1,Workbook2, Array(False))

	End Function

	''' <summary>
    ''' Compare 2 Excel sheets cell by cell with ignoring the Header row
    ''' </summary>
    ''' <param name="Workbook1" type="string">Location to the Excel WorkBook 1</param>
    ''' <param name="Workbook2" type="string">Location to the Excel WorkBook 2</param>
    ''' <returns>True/False</returns>
	Public Function Compare(Byval Workbook1, Byval Workbook2)

		Compare = CompareExcel(Workbook1,Workbook2, Array(True))

	End Function


	''' <summary>
    ''' Compare 2 Excel sheets cell by cell with ignoring the Header row
    ''' </summary>
    ''' <param name="Workbook1" type="string">Location to the Excel WorkBook 1</param>
    ''' <param name="Workbook2" type="string">Location to the Excel WorkBook 2</param>
	''' <param name="arr" type="array">determine only compare special columns</param>
    ''' <returns>True/False</returns>
	Public Function CompareSpecifiedColumns(Byval Workbook1, Byval Workbook2, ByVal arr)

		CompareSpecifiedColumns = CompareExcel(Workbook1,Workbook2, Array(True,arr))

	End Function

	''' <summary>
    ''' Compare 2 Excel sheets cell by cell with ignoring the Header row
    ''' </summary>
    ''' <param name="Workbook1" type="string">Location to the Excel WorkBook 1</param>
    ''' <param name="Workbook2" type="string">Location to the Excel WorkBook 2</param>
	''' <param name="arr" type="array">determine only compare special columns</param>
    ''' <returns>True/False</returns>
	Public Function CompareSpecifiedColumnsIncludingHeaderRow(Byval Workbook1, Byval Workbook2, ByVal arr)

		CompareSpecifiedColumnsIncludingHeaderRow = CompareExcel(Workbook1,Workbook2, Array(false,arr))

	End Function

	''' <summary>
    ''' Compare 2 Excel sheets cell by cell
    ''' </summary>
    ''' <param name="Workbook1" type="string">Location to the Excel WorkBook 1</param>
    ''' <param name="Workbook2" type="string">Location to the Excel WorkBook 2</param>
	''' <param name="arr" type="array">determine if ignore Header row and only compare special columns</param>
    ''' <returns>True/False</returns>
	Public Function CompareExcel(Byval Workbook1, Byval Workbook2, ByVal arr)
	
		Dim objWorkbook1,objWorkbook2,objWorksheet1,objWorksheet2
		
		Set objWorkbook1= xlsApp.Workbooks.Open(Workbook1)
		Set objWorkbook2= xlsApp.Workbooks.Open(Workbook2)
		
		Set objWorksheet1 = objWorkbook1.Worksheets(1)
		Set objWorksheet2 = objWorkbook2.Worksheets(1)
		
		Dim GetUsedRowCount1,GetUsedRowCount2,GetUsedColumnCount1,GetUsedColumnCount2,GetUsedMaxRowCount,GetUsedMaxColumnCount
		'Const xlUp = -4162
		'Const xlToRight = -4161
		'GetUsedRowCount1 = objWorksheet1.Range("A65536").End(xlUp).Row
		'GetUsedRowCount2 = objWorksheet2.Range("A65536").End(xlUp).Row
		'GetUsedRowCount1 = objWorksheet1.Cells(1,1).End(xlUp).Row
		'GetUsedRowCount2 = objWorksheet2.Cells(1,1).End(xlUp).Row	
		'GetUsedColumnCount1 = objWorksheet1.Cells(1, 1).End(xlToRight).Column
		'GetUsedColumnCount2 = objWorksheet2.Cells(1, 1).End(xlToRight).Column
		
		GetUsedRowCount1 = objWorksheet1.UsedRange.Rows.Count
		GetUsedRowCount2 = objWorksheet2.UsedRange.Rows.Count
		GetUsedColumnCount1 = objWorksheet1.UsedRange.Columns.Count
		GetUsedColumnCount2 = objWorksheet2.UsedRange.Columns.Count
		
		If(GetUsedRowCount1 < GetUsedRowCount2) Then
			GetUsedMaxRowCount = GetUsedRowCount2
		Else
			GetUsedMaxRowCount = GetUsedRowCount1
		End If
		
		If(GetUsedColumnCount1 < GetUsedColumnCount2) Then
			GetUsedMaxColumnCount = GetUsedColumnCount2
		Else
			GetUsedMaxColumnCount = GetUsedColumnCount1
		End If	
		
		Dim FirstCell,LastCell,cell
		FirstCell = objWorksheet1.Cells(1,1).Address(0,0)
		LastCell = objWorksheet1.Cells(GetUsedMaxRowCount,GetUsedMaxColumnCount).Address(0,0)
		
		Dim WorkSheet12DArray, WorkSheet22DArray
		Dim row, col, firstrow, bIgnoreHeaderRow, bCompareAllColumns
		Dim iCountDifferent, minus
		WorkSheet12DArray = objWorksheet1.Range(FirstCell & ":" & LastCell).Value
		WorkSheet22DArray = objWorksheet2.Range(FirstCell & ":" & LastCell).Value
		
		Select Case UBound(arr)
			Case 0
				bIgnoreHeaderRow = arr(0)
				If bIgnoreHeaderRow Then
					firstrow = 2
				else
					firstrow = 1
				End If
				For row = firstrow To GetUsedMaxRowCount
					For col = 1 To GetUsedMaxColumnCount
						If Trim(UCase(WorkSheet12DArray(row, col))) <> Trim(UCase(WorkSheet22DArray(row, col))) Then
							If IsNumeric(WorkSheet12DArray(row, col)) And IsNumeric(WorkSheet22DArray(row, col)) Then
								minus = WorkSheet22DArray(row, col) - WorkSheet12DArray(row, col)
								'if one value is 11, other value is 11.000, then the difference should be ignored
								If minus <> 0 Then
									'Color Palette and the 56 Excel ColorIndex Colors http://www.mvps.org/dmcritchie/excel/colors.htm
									objWorksheet1.Cells(row, col).Interior.ColorIndex = 22
									iCountDifferent = iCountDifferent + 1
								End If
							ElseIf IsDate(WorkSheet12DArray(row, col)) And IsDate(WorkSheet22DArray(row, col)) Then
								If DateDiff("s",WorkSheet12DArray(row, col),WorkSheet22DArray(row, col)) <> 0 Then
									objWorksheet1.Cells(row, col).Interior.ColorIndex = 22
									iCountDifferent = iCountDifferent + 1
								End If
							else	
								'Color Palette and the 56 Excel ColorIndex Colors http://www.mvps.org/dmcritchie/excel/colors.htm
								objWorksheet1.Cells(row, col).Interior.ColorIndex = 22
								iCountDifferent = iCountDifferent + 1
							End if
'						Else
'							'objWorksheet1.Cells(row, col).Interior.ColorIndex = 0
						End If
					Next
				next		
			Case 1
				bIgnoreHeaderRow = arr(0)
				If bIgnoreHeaderRow Then
					firstrow = 2
				else
					firstrow = 1
				End If
				For row = firstrow To GetUsedMaxRowCount
					For col = 1 To GetUsedMaxColumnCount
						If CheckValueExistsinArray(col, arr(1)) then
							If Trim(UCase(WorkSheet12DArray(row, col))) <> Trim(UCase(WorkSheet22DArray(row, col))) Then
								If IsNumeric(WorkSheet12DArray(row, col)) And IsNumeric(WorkSheet22DArray(row, col)) Then
									minus = WorkSheet22DArray(row, col) - WorkSheet12DArray(row, col)
									'if one value is 11, other value is 11.000, then the difference should be ignored
									If minus <> 0 Then
										'Color Palette and the 56 Excel ColorIndex Colors http://www.mvps.org/dmcritchie/excel/colors.htm
										objWorksheet1.Cells(row, col).Interior.ColorIndex = 22
										iCountDifferent = iCountDifferent + 1
									End If
								ElseIf IsDate(WorkSheet12DArray(row, col)) And IsDate(WorkSheet22DArray(row, col)) Then
									If DateDiff("s",WorkSheet12DArray(row, col),WorkSheet22DArray(row, col)) <> 0 Then
										objWorksheet1.Cells(row, col).Interior.ColorIndex = 22
										iCountDifferent = iCountDifferent + 1
									End If
								else	
									'Color Palette and the 56 Excel ColorIndex Colors http://www.mvps.org/dmcritchie/excel/colors.htm
									objWorksheet1.Cells(row, col).Interior.ColorIndex = 22
									iCountDifferent = iCountDifferent + 1
								End if
							End If
						End if	
					Next
				Next

		End Select

		If iCountDifferent <> 0 Then
			CompareExcel = False
		Else
			CompareExcel = true	
		End If
			
		objWorkbook1.Save
		objWorkbook2.Save
		objWorkbook1.Close
		objWorkbook2.Close
		Set objWorkbook1 = Nothing
		Set objWorkbook2 = Nothing
		Set objWorksheet1 = Nothing
		Set objWorksheet2 = Nothing
	
	End Function
	
	''' <summary>
    ''' Check whether a value exists in  a given array
    ''' </summary>
    ''' <param name="str" type="string">The str to be checked</param>
    ''' <param name="arr" type="array">The array source</param>
    ''' <return>true/false</return>
    ''' <remarks></remarks>
	Public Function CheckValueExistsinArray(ByVal str, ByVal arr)
		
		CheckValueExistsinArray=False
		Dim i
		For i = LBound(arr) to UBound(arr)
			If LCase(Trim(arr(i))) = LCase(Trim(str)) then
				CheckValueExistsinArray = true
				Exit Function
			end if
		Next
	    
	End Function

End Class

Public Function CompareTwoExcelLib
	
	Set CompareTwoExcelLib = New ClsCompareTwoExcelLib
	
End Function




