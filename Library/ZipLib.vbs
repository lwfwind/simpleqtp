Option Explicit

''' #########################################################
''' <summary>
''' A Library to zip/unzip file
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Class ClsZipLib

	''' <summary>
    ''' Zip file using windows shell
    ''' </summary>
    ''' <param name="sFile" type="string">the location of the file to be zip</param>
    ''' <param name="sZipFile" type="string">the location of the generated zip file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <example>
	''' Call Compress ("c:\Config.xls", "c:\Config.zip")
    ''' </example>
	Public Function Compress(ByVal sFile,ByVal sZipFile)
	
		Set oZipShell = CreateObject("WScript.Shell")
		Set oZipFSO = CreateObject("Scripting.FileSystemObject")
		
		If Not oZipFSO.FileExists(sZipFile) Then
		NewZip(sZipFile)
		End If
		
		Set oZipApp = CreateObject("Shell.Application")
		
		sZipFileCount = oZipApp.NameSpace(sZipFile).items.Count
		
		aFileName = Split(sFile, "\")
		sFileName = (aFileName(Ubound(aFileName)))
		
		'listfiles
		sDupe = False
		For Each sFileNameInZip In oZipApp.NameSpace(sZipFile).items
		If LCase(sFileName) = LCase(sFileNameInZip) Then
		sDupe = True
		Exit For
		End If
		Next
		
		If Not sDupe Then
		oZipApp.NameSpace(sZipFile).Copyhere sFile
		
		'Keep script waiting until Compressing is done
		On Error Resume Next
		sLoop = 0
		Do Until sZipFileCount < oZipApp.NameSpace(sZipFile).Items.Count
		wait 0.1
		sLoop = sLoop + 1
		Loop
		On Error GoTo 0
		End If
		
	End Function
	
	Public Sub NewZip(sNewZip)
		
		Set oNewZipFSO = CreateObject("Scripting.FileSystemObject")
		Set oNewZipFile = oNewZipFSO.CreateTextFile(sNewZip)
		
		oNewZipFile.Write Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
		
		oNewZipFile.Close
		Set oNewZipFSO = Nothing
		
		wait 1
		
	End Sub
	
	''' <summary>
    ''' Extract all files and folders from a compressed file (ZIP, CAB, etc.) using windows shell
    ''' </summary>
    ''' <param name="sZipFile" type="string">the fully qualified path of the ZIP file</param>
    ''' <param name="sTargetDir" type="string">the fully qualified path of the (existing) destination folder</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <example>
	''' Extract "C:\Config.zip", "C:\"
	''' Extract "C:\test.cab", "C:\"
	''' Extract "C:\test2", "C:\test3"
    ''' </example>
	Sub Extract(ByVal sZipFile, ByVal sTargetDir)
	
	    Dim intOptions, objShell, objSource, objTarget
	
	    ' Create the required Shell objects
	    Set objShell = CreateObject( "Shell.Application" )
	
	    ' Create a reference to the files and folders in the ZIP file
	    Set objSource = objShell.NameSpace( sZipFile ).Items( )
	
	    ' Create a reference to the target folder
	    Set objTarget = objShell.NameSpace( sTargetDir )
	
	    ' These are the available CopyHere options, according to MSDN
	    ' (http://msdn2.microsoft.com/en-us/library/ms723207.aspx).
	    ' On my test systems, however, the options were completely ignored.
	    '      4: Do not display a progress dialog box.
	    '      8: Give the file a new name in a move, copy, or rename
	    '         operation if a file with the target name already exists.
	    '     16: Click "Yes to All" in any dialog box that is displayed.
	    '     64: Preserve undo information, if possible.
	    '    128: Perform the operation on files only if a wildcard file
	    '         name (*.*) is specified.
	    '    256: Display a progress dialog box but do not show the file
	    '         names.
	    '    512: Do not confirm the creation of a new directory if the
	    '         operation requires one to be created.
	    '   1024: Do not display a user interface if an error occurs.
	    '   4096: Only operate in the local directory.
	    '         Don't operate recursively into subdirectories.
	    '   8192: Do not copy connected files as a group.
	    '         Only copy the specified files.
	    intOptions = 256
	
	    ' UnZIP the files
	    objTarget.CopyHere objSource, intOptions
	
	    ' Release the objects
	    Set objSource = Nothing
	    Set objTarget = Nothing
	    Set objShell  = Nothing
	    
	End Sub 

End Class

Public Function ZipLib()

	Set ZipLib = New ClsZipLib

End Function	
