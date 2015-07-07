Option Explicit

''' ##################################################################
''' <summary>
''' Schedule a task using windows inner Scheduled Task to run one QTP Test automatically
''' </summary>
''' <remarks>
''' Please make sure putting this VBS file into same directory with AutoRun.VBS
''' Double-click this file will schedule one QTP Test automatically
''' </remarks>
''' ##################################################################

Class ClsScheduleAOM
	
	''' <summary>
    ''' Disable windows 'ScreenSaver'
    ''' </summary>
    ''' <remarks></remarks>
	Public Function DisableScreenSaver()
	
		Dim WshShell  
		Set WshShell = WScript.CreateObject("WScript.Shell")  
		WshShell.RegWrite "HKCU\Control Panel\Desktop\ScreenSaveActive",0,"REG_SZ" 
		Set WshShell = Nothing
	
	End Function
	
	''' <summary>
    ''' Delete specific task using windows inner Scheduled Task
    ''' </summary>
    ''' <remarks></remarks>
	Public Function DeleteTask(ByVal sTaskName)
	
		Dim WshShell, DeleteParemeters
		Set WshShell = CreateObject("WScript.Shell")
		DeleteParemeters = "/Delete /tn " & Chr(34) & sTaskName & Chr(34) & " /f"
		WshShell.Run "schtasks.exe " & DeleteParemeters
		Set WshShell = nothing
	End Function

	''' <summary>
    ''' Add specific task using windows inner Scheduled Task
    ''' </summary>
    ''' <param name="sTaskName" type="string">Specifies a name for the task</param>
    ''' <param name="sStartDay" type="string">Specifies the day that the task starts in MM/DD/YYYY format</param>
    ''' <param name="sStartTime" type="string">Specifies the time of day that the task starts in HH:MM:SS 24-hour format</param>
    ''' <param name="sSchedule" type="string">Specifies the schedule type. Valid values are MINUTE, HOURLY, DAILY, WEEKLY, MONTHLY, ONCE, ONSTART, ONLOGON, ONIDLE</param>
    ''' <remarks>Detail for schtasks.exe please refer to http://www.microsoft.com/resources/documentation/windows/xp/all/proddocs/en-us/schtasks.mspx?mfr=true</remarks>
	Public Function AddTask(ByVal sTaskName, ByVal sStartDay, ByVal sStartTime, ByVal sSchedule)
	
		Dim WshShell, sScriptLocation, AddParemeters
		Set WshShell = CreateObject("WScript.Shell")
		Call DisableScreenSaver
		Call DeleteTask(sTaskName)
		'A Scheduled Task Does Not Run When You Use Schtasks.exe to Create It 
		'and When the Path of the Scheduled Task Contains a Space
		'Please refer to http://support.microsoft.com/kb/823093/en-us
		sScriptLocation = "\" & Chr(34) & WshShell.CurrentDirectory & "\AutoRun.vbs" & "\" & Chr(34)
		AddParemeters = "/create /ru system /tn " & Chr(34) & sTaskName & Chr(34) & " /tr " & Chr(34) & sScriptLocation & Chr(34) & " /sc " & sSchedule & " /sd " & sStartDay &  " /st " & sStartTime
		WshShell.Run "schtasks.exe " & AddParemeters
		Set WshShell = Nothing
		
	End Function

	Public Function UserInput(ByVal myPrompt)
		' This function prompts the user for some input.
		' When the script runs in CSCRIPT.EXE, StdIn is used,
		' otherwise the VBScript InputBox( ) function is used.
		' myPrompt is the the text used to prompt the user for input.
		' The function returns the input typed either on StdIn or in InputBox( ).
	    ' Check if the script runs in CSCRIPT.EXE
	    If UCase( Right( WScript.FullName, 12 ) ) = "\CSCRIPT.EXE" Then
	        ' If so, use StdIn and StdOut
	        WScript.StdOut.Write myPrompt & " "
	        UserInput = WScript.StdIn.ReadLine
	    Else
	        ' If not, use InputBox( )
	        UserInput = InputBox(myPrompt)
	    End If
	End Function

End Class


Dim ScheduleAOM, sTaskName, sStartDay, sStartTime, sTime, sSchedule, arr
Dim WshShell
Dim oExcel, oBook, oModule
Dim strRegKey, strCode, x, y, toTime, sFlag
Set ScheduleAOM = New ClsScheduleAOM
sTime = ScheduleAOM.UserInput( "Please Enter the scheduled time in MM/DD/YYYY HH:MM:SS format")
Do While Len(sTime) <> 19
	sTime = ScheduleAOM.UserInput( "The entered format is not valid, please re-enter")
Loop
arr = Split(sTime, " ")
sStartDay = arr(0)
sStartTime = arr(1)
sSchedule = "once"
sTaskName = "Schedule running QTP"
ScheduleAOM.AddTask sTaskName, sStartDay, sStartTime, sSchedule

'Move the cursor every 5 minutes util reach the scheduled time
set WshShell = CreateObject("wscript.Shell")
strRegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\$\Excel\Security\AccessVBOM"
strRegKey = Replace(strRegKey, "$", oExcel.Version)
WshShell.RegWrite strRegKey, 1, "REG_DWORD"
Set oExcel = CreateObject("Excel.Application") '创建 Excel 对象
Set oBook = oExcel.Workbooks.Add '添加工作簿
Set oModule = obook.VBProject.VBComponents.Add(1) '添加模块
strCode = _
"Private Type POINTAPI : X As Long : Y As Long : End Type"                                                                                                                            & vbCrLf & _
"Private Declare Function SetCursorPos Lib ""user32"" (ByVal x As Long, ByVal y As Long) As Long"                                                                                     & vbCrLf & _
"Private Declare Function GetCursorPos Lib ""user32"" (lpPoint As POINTAPI) As Long"                                                                                                  & vbCrLf & _
"Private Declare Sub mouse_event Lib ""user32"" Alias ""mouse_event"" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)" & vbCrLf & _
"Public Function GetXCursorPos() As Long"                                                                                                                                             & vbCrLf & _
    "Dim pt As POINTAPI : GetCursorPos pt : GetXCursorPos = pt.X"                                                                                                                     & vbCrLf & _
"End Function"                                                                                                                                                                        & vbCrLf & _
"Public Function GetYCursorPos() As Long"                                                                                                                                             & vbCrLf & _
    "Dim pt As POINTAPI: GetCursorPos pt : GetYCursorPos = pt.Y"                                                                                                                      & vbCrLf & _
"End Function"
oModule.CodeModule.AddFromString strCode '在模块中添加 VBA 代码

sFlag = true
toTime = CDate(sTime)
Do While CInt(DateDiff("n",Now,toTime)) > 0
	x = oExcel.Run("GetXCursorPos") '获取鼠标 X 坐标
	y = oExcel.Run("GetYCursorPos") '获取鼠标 Y 坐标
	If sFlag Then
		oExcel.Run "SetCursorPos", x, y '设置鼠标 X Y 坐标
	Else
		oExcel.Run "SetCursorPos", x+1, y
	End if
	WScript.Sleep 300000
	If sFlag Then
		sFlag = false
	Else
		sFlag = true
	End if
Loop

Set oModule = Nothing
Set oBook = Nothing
Set oExcel = Nothing
Set WshShell = nothing

