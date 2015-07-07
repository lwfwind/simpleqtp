Option Explicit

'''#######################################################################
''' <summary>
''' Centralize all constant definitions
''' </summary>
''' <remarks></remarks>
''' <contents>
	''' <path>
	''' DriverPath
	''' LibraryPath
	''' ScriptPath
	''' DataPath
	''' TestSuitesPath
	''' TestCasesPath
	''' ConfigPath
	''' ORPath
	''' ReportPath
	''' Report_Screenshot_Path
	''' Report_LinkToFile_Path
	''' TempPath
	''' </path>
	
	''' <config entries>
	''' PreExecutionSetupFile
	''' </config entries>
	
	''' <reports entries>
	''' ReportLib
	''' </reports entries>
	
	''' <ORPath entries>
	''' GlobalObjectsMapToOR
	''' </ORPath entries>
	
	''' <library entries>
	''' ExcelLib
	''' XlsDictLib
	''' CompareTwoExcelLib.vbs
	''' ArrayLib
	''' StringLib
	''' WordLib
	''' DateTimeLib
	''' DBLib
	''' EmailLib
	''' FSOLib
	''' FTPLib
	''' LoggerLib
	''' ObjectGeneralLib
	''' RegLib
	''' ZipLib
	''' </library entries>
	
''' </contents>
'''#######################################################################

Class GlobalConstants

	Private oFSO
	Private BasePath

	''' <summary>
    ''' Class Initialization procedure. Initialize the BasePath variable
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()

		Dim GetParentFolderPath
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		GetParentFolderPath = oFSO.GetParentFolderName(PathFinder.CurrentTestPath)
		BasePath =  GetParentFolderPath & "\"
	   	
	End Sub

	''' <summary>
    ''' Class_Terminate procedure. Clear the BasePath variable
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
		
		BasePath = ""
		Set oFSO = nothing
	   	
	End Sub

	''' #######################################################################
	''' Path Constants
	''' #######################################################################

	''' <summary>
    ''' Driver Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get DriverPath()
	
	   DriverPath = BasePath & "Driver\"
	   
	End Property

	''' <summary>
    ''' Library Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get LibraryPath()
	
	   LibraryPath = BasePath & "Library\"
	   
	End Property
	
	''' <summary>
    ''' Script Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ScriptPath()
	
	   ScriptPath = BasePath & "Scripts\"
	   
	End Property

	''' <summary>
    ''' Data Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get DataPath()
	
	   DataPath = BasePath & "DataFiles\"
	   
	End Property
	
	''' <summary>
    ''' Data TestSuites Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get TestSuitesPath()
	
	   TestSuitesPath = DataPath & "TestSuites\"
	   
	End Property

	''' <summary>
    ''' Data TestCases Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get TestCasesPath()
	
	   TestCasesPath = DataPath & "TestCases\"
	   
	End Property
	
	''' <summary>
    ''' Config Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ConfigPath()
	
	   ConfigPath = BasePath & "Config\"
	   
	End Property

	''' <summary>
    ''' Object Repository Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ORPath()
	
	   ORPath = BasePath & "ORFiles\"
	   
	End Property
	
	''' <summary>
    ''' Report Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ReportPath()
	
	   ReportPath = BasePath & "Reports\"
	   
	End Property
	
	''' <summary>
    ''' Report Screenshot Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get Report_Screenshot_Path()
	
	   Report_Screenshot_Path = BasePath & "Reports\_Screenshots\"
	   
	End Property
	
	''' <summary>
    ''' Report Link To File Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get Report_LinkToFile_Path()
	
	   Report_LinkToFile_Path = BasePath & "Reports\_filepath\"
	   
	End Property

	''' <summary>
    ''' Template Path
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get TempPath()
	
	   TemplatePath = BasePath & "Temp\"
	   
	End Property
	
	''' #######################################################################
	''' Config Entries
	''' #######################################################################
	
	''' <summary>
    ''' PreExecutionSetup File
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get PreExecutionSetupFile()
	   
	   PreExecutionSetupFile = ConfigPath & "PreExecutionSetup.xls"
	   
	End Property
	
	''' #######################################################################
	''' Reports Entries
	''' #######################################################################
	
	''' <summary>
    ''' Report Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ReportLib()
	   
	   ReportLib = ReportPath & "ReportLib.vbs"
	   
	End Property
	
	''' #######################################################################
	''' ORPath Entries
	''' #######################################################################
		
	''' <summary>
    ''' Flight OR file
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get FlightOR()
	   
	   FlightOR = ORPath & "Flight.tsr"
	   
	End Property
	
	''' <summary>
    ''' GlobalObjectsMapToOR file
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get GlobalObjectsMapToOR()
	   
	   GlobalObjectsMapToOR = ORPath & "GlobalObjectsMapToOR.vbs"
	   
	End Property

	''' #######################################################################
	''' Library Entries
	''' #######################################################################
	
	''' <summary>
    ''' Array Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ArrayLib()
	   
	   ArrayLib = LibraryPath & "ArrayLib.vbs"
	   
	End Property
	
	''' <summary>
    ''' String Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get StringLib()
	   
	   StringLib = LibraryPath & "StringLib.vbs"
	   
	End Property

	''' <summary>
    ''' Excel Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ExcelLib()
	   
	   ExcelLib = LibraryPath & "ExcelLib.vbs"
	   
	End Property
	
	''' <summary>
    ''' CompareTwoExcel Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get CompareTwoExcelLib()
	   
	   CompareTwoExcelLib = LibraryPath & "CompareTwoExcelLib.vbs"
	   
	End Property
	
	''' <summary>
    ''' FileSystemObject Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get FSOLib()
	   
	   FSOLib = LibraryPath & "FSOLib.vbs"
	   
	End Property

	''' <summary>
    ''' Dictionary based on excel Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get XlsDictLib()
	   
	   XlsDictLib = LibraryPath & "XlsDictLib.vbs"
	   
	End Property
		
	''' <summary>
    ''' Reg Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get RegLib()
	   
	   RegLib = LibraryPath & "RegLib.vbs"
	   
	End Property

	''' <summary>
    ''' Word Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get WordLib()
	   
	   WordLib = LibraryPath & "WordLib.vbs"
	   
	End Property

	''' <summary>
    ''' Zip Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ZipLib()
	   
	   ZipLib = LibraryPath & "ZipLib.vbs"
	   
	End Property	

	''' <summary>
    ''' DB Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get DBLib()
	   
	   DBLib = LibraryPath & "DBLib.vbs"
	   
	End Property	
	
	''' <summary>
    ''' Email Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get EmailLib()
	   
	   EmailLib = LibraryPath & "EmailLib.vbs"
	   
	End Property	
	
	''' <summary>
    ''' DateTime Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get DateTimeLib()
	   
	   DateTimeLib = LibraryPath & "DateTimeLib.vbs"
	   
	End Property	
	
	''' <summary>
    ''' FTP Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get FTPLib()
	   
	   FTPLib = LibraryPath & "FTPLib.vbs"
	   
	End Property

	''' <summary>
    ''' FTP Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get LoggerLib()
	   
	   LoggerLib = LibraryPath & "LoggerLib.vbs"
	   
	End Property

	''' <summary>
    ''' ObjectGeneral Lib
    ''' </summary>
    ''' <remarks></remarks>
	Public Property Get ObjectGeneralLib()
	   
	   ObjectGeneralLib = LibraryPath & "ObjectGeneralLib.vbs"
	   
	End Property
	
End Class

Public Function GlobalConst()
	
	Set GlobalConst = New GlobalConstants

End Function
