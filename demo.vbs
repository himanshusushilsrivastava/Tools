'###########################################################################
'    Created    :
'    By         : UHOne Facets
'    Date       : 07/11/2018
'    Purpose    : This script fetches all the input file names from nas drive
'                 and stores them in the app server ftp path
'---------------------------------------------------------------------------
'    Arguments  : 1. Path of Source File Directory
'                 2. Path of Destination Files Directory
'---------------------------------------------------------------------------
'###########################################################################
'    Created by   Remarks                                        Date
'---------------------------------------------------------------------------
'    CTS          Initial Version                                03/09/2018
'---------------------------------------------------------------------------
'###########################################################################

Option Explicit

Call CopySourceToDestFile

Private Sub CopySourceToDestFile()

Dim oFSO
Dim l_wshShell
Dim l_fsoFileSysObj
Dim l_argValues
Dim l_strSrcFileDir
Dim l_strDestFileDir
Dim l_strFileName
Dim l_strFileExt
Dim l_strFileRename
Dim l_strFileType
Dim l_intArgsCount

Dim l_txtStrOutput
Dim l_fdInputFolder
Dim l_flInputFile
Dim l_strDestinationFile
Dim l_intFileCount

    'On Error Resume Next
    l_intFileCount = 0

    ' Creating a Shell script Object, to write the error
    Set l_wshShell = wscript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Non Critical Error : " & _
            "Creation of Shell Object Failed; Err Msg : " & Err.Description
        wscript.Quit (99)
    End If

    ' Creating a FileSystem Object
    Set l_fsoFileSysObj = wscript.CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Non Critical Error : " & _
            "Creation of FileSystem Object Failed; Err Msg : " & Err.Description
        wscript.Quit (99)
    End If

    ' Get the Command line arguments in the argument object
    Set l_argValues = wscript.Arguments
    Set oFSO = CreateObject("Scripting.FileSystemObject") 

l_intArgsCount = l_argValues.Count
If l_intArgsCount <> 2 Then
	        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Critical Error : " & _
            "Wrong Number of Arguments. Only 2 arguments are allowed. 1.Source Directory 2.Destination Directory"
        Set l_wshShell = Nothing 
        wscript.Quit (3)
End If	

If l_fsoFileSysObj.FolderExists(l_argValues.Item(0)) Then
	l_strSrcFileDir     = l_argValues.Item(0)
Else 
		        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Critical Error : " & _
            "Args 1,Source Directory is not a valid path"
        Set l_wshShell = Nothing 
        wscript.Quit (3)
End If

If l_fsoFileSysObj.FolderExists(l_argValues.Item(1)) Then
	l_strDestFileDir    = l_argValues.Item(1)
Else 
			        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Critical Error : " & _
            "Args 2,Destination Directory is not a valid path"
        Set l_wshShell = Nothing 
        wscript.Quit (3)	
End If	

    l_strDestFileDir = CheckSlash(l_strDestFileDir)
   

    Set l_fdInputFolder = l_fsoFileSysObj.GetFolder(l_strSrcFileDir)
    
		For Each l_flInputFile In l_fdInputFolder.Files
		l_strFileName = l_fsoFileSysObj.GetBaseName(l_flInputFile)
		l_strFileExt  = l_fsoFileSysObj.GetExtensionName(l_flInputFile)		
		l_strFileType = Left(Right(l_strFileName,4),1)
		
				 l_strFileName = l_fsoFileSysObj.GetFileName(l_flInputFile)
         If l_strFileExt = "999" Then
         	If l_strFileType = "P" Then
            l_strFileRename = l_strDestFileDir + "5010_999.GRIFAC.Multiplan.837P." + Right("00" & Year(Now()), 4) + "-" + Right("00" & Month(Now()), 2) + "-" +Right("00" & Day(Now()), 2) + "T" + Right("00" & Hour(Now()), 2) + Right("00" & Minute(Now()), 2) + Right("00" & Second(Now()), 2) +"." + l_strFileExt
          ElseIf l_strFileType = "I" Then  
          	l_strFileRename = l_strDestFileDir + "5010_999.GRIFAC.Multiplan.837I." + Right("00" & Year(Now()), 4) + "-" + Right("00" & Month(Now()), 2) + "-" + Right("00" & Day(Now()), 2) + "T" + Right("00" & Hour(Now()), 2) + Right("00" & Minute(Now()), 2) + Right("00" & Second(Now()), 2) +"." + l_strFileExt
          End If	            
         else
         	If l_strFileType = "P" Then
            l_strFileRename = l_strDestFileDir + "5010_GRICFAC_HCI" + Right("00" & Year(Now()), 4) + Right("00" & Month(Now()), 2) + Right("00" & Day(Now()), 2) + Right("00" & Hour(Now()), 2) + Right("00" & Minute(Now()), 2) + Right("00" & Second(Now()), 2) +"." + "txt"
          ElseIf l_strFileType = "I" Then  
          	l_strFileRename = l_strDestFileDir + "5010_GRICFAC_UCI" + Right("00" & Year(Now()), 4) + Right("00" & Month(Now()), 2) + Right("00" & Day(Now()), 2) + Right("00" & Hour(Now()), 2) + Right("00" & Minute(Now()), 2) + Right("00" & Second(Now()), 2) +"." + "txt"
          End If	
         End If
            oFSO.MoveFile l_flInputFile, l_strFileRename
            l_intFileCount = l_intFileCount + 1
            WScript.Sleep 1000 'Delay
    Next  

    Set l_flInputFile   = Nothing
    Set l_fdInputFolder = Nothing
    Set l_txtStrOutput  = Nothing
    Set l_fsoFileSysObj = Nothing
    Set l_argValues     = Nothing

    If l_intFileCount = 0 Then

        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Critical Error : " & _
            "No Input Files are available for processing"
        Set l_wshShell = Nothing 
        wscript.Quit (3)
    Else 
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - File Count Copied : " & l_intFileCount 
    End If
    
    Set l_wshShell = Nothing
    
    ' Checking for run time errors
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Run Time Error : " & _
            "Err Msg : " & Err.Description
       
        Set l_flInputFile   = Nothing
        Set l_fdInputFolder = Nothing
        Set l_txtStrOutput  = Nothing
        Set l_fsoFileSysObj = Nothing
        Set l_argValues     = Nothing
        Set l_wshShell      = Nothing
    End If

End Sub

'###########################################################################
'Function Name     : CheckSlash
'Purpose           : To append backward slash ("\") to the folder name if its
'                    last character is not "\"
'Input             :
'                    p_strFoldName - Folder path
'Output            : Folder path with "\" appended
'###########################################################################
Private Function CheckSlash(ByVal p_strFoldName)

    On Error Resume Next

    ' Checking for the presence of "\" as the last character in string containing the folder path
    If ((Right(Trim(p_strFoldName), 1) <> "\") and Len(Trim(p_strFoldName)) <> 0) Then
        CheckSlash = p_strFoldName & "\"
    Else
        CheckSlash = p_strFoldName
    End If
        
End Function
