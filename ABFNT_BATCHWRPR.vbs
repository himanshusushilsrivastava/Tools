'###########################################################################
'----------------------------- NT BATCH WRAPPER -----------------------------
'    Created:
'    By         : UHGIS offshore
'    Date       : 03/06/2017
'    Purpose    : To read the configuration xml,perform any pre/post processes,
'                 modify the run book and execute the batch.
'---------------------------------------------------------------------------
'    Description: This script receives batch name and database name as
'                 parameters.
'
'                 Based on the batch name, the corresponding section in the
'                 configuration file is read and all the batch specific information
'                 are fetched and stored in a global dictionary object. Also, the
'                 generic information are fetched and stored in the dictionary object.
'
'                 The various steps in the pre process xml are read and the
'                 required steps are executed.
'
'                 The runbook for the batch is modified with the override parameters
'                 and the modified runbook is run to execute the batch.
'
'                 The various steps in the post process xml are read and the
'                 required steps are executed.
'---------------------------------------------------------------------------
'    Arguments  : 1. Batch Name
'                 2. Database Name
'---------------------------------------------------------------------------
'    Assumptions: 1. The configuration file "commonsystemprop.xml" is available
'                    in the application path.
'                 2. The VB DLL JD_NTBATCH_CMNFUNCS is registered in the
'                    system registry.
'                 3. The folders configured in the configuration file
'                    "commonsystemprop.xml" are available.
'                 4. The executable Blat.exe used to send mail is available
'                    in the path mentioned in the commonsystemprop.xml.
'                 5. The input files for the batches are named according to
'                    the naming convention mentioned in the configuration file.
'###########################################################################
'    Created by   Remarks                                        Date
'---------------------------------------------------------------------------
'    UHGIS        Initial Version                                03/06/2017
'---------------------------------------------------------------------------
'###########################################################################

Option Explicit

' Global variable declarations
Dim g_dicBatchInfo          'Dictionary object to store generic config info and batch specific info
Dim g_dicOverrideInfo       'Dictionary object to store override parameters for a batch
'Dim g_dllCmnFunctions       'jdh_cmnfuncs class object used to access the functions in the JD_NTBATCH_CMNFUNCS.dll
Dim g_arrPreBypassInfo      'Array to store any bypass steps or bypass actions for the pre-process
Dim g_arrPostBypassInfo     'Array to store any bypass steps or bypass actions for the post-process
Dim g_loadXml               'XML DOM object to load an xml
Dim g_loadDBXml               'XML DOM object to load an xml
Dim g_connDatabase          'ADODB Database Connection Object to connect to the Database

Const SYSTEM_FILE_PATH = "E:\TriZetto\Facets\Regions\"

' Calling the NT Batch Wrapper Module
Call main

'###########################################################################
'Subroutine Name   : main
'Purpose           : 1. To call the f_ValidateInputParameter function to
'                       validate the command line arguments
'                    2. To call the f_ReadConfigFile and f_ReadSystemFile
'                       functions to fetch the information from the config
'                       file and system default xml.
'                    3. To call the f_ExecuteProcessFile to execute the pre/
'                       post processes for a batch.
'                    4. To call the f_BuildRunbook function to modify the
'                       run book with the override parameters.
'                    5. To call the f_ExecuteBatch/f_Execute837Batch functions
'                       to execute the respective batches with the modified
'                       run book.
'Input             : None
'Output            : None
'###########################################################################
Private Sub main()

    On Error Resume Next

    Dim l_strProvider
    Dim l_strUsername
    Dim l_strPassword
    Dim l_strDBname
    Dim l_strCustDBname 
    Dim l_strServerName
    Dim l_strServerPort

    Dim l_strAdditionalParam
    Dim l_wshShell                     ' Shell object to write the Eventlog
    
	
	Dim l_intFileCount
	Dim l_fsoFileSysObj
	Dim oFSO
    Dim l_flInputFile
    Dim l_flBadFiles
	Dim l_fdInputFolder
	
		
    ' Creating a Shell script Object, to write the error
    Set l_wshShell = wscript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
            "Creation of Shell Object Failed; Err Msg : " & Err.Description
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Creation of Shell Object Failed; Err Msg : " & Err.Description
        wscript.Quit (99)
    End If

    ' Creating the object to access the JD_NTBATCH_CMNFUNCS DLL functions
 ''   Set g_dllCmnFunctions = WScript.CreateObject("JD_NTBATCH_CMNFUNCS.jdh_cmnfuncs")
''    If Err.Number <> 0 Then
        ''' Writing the Error message in the Event log
''        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
''            "Creation of Custom DLL Failed; Err Msg : " & Err.Description
        ''wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Creation of Custom DLL Failed; Err Msg : " & Err.Description
''        wscript.Quit (99)
    ''End If

    ' Create the XML DOM object
    Set g_loadXml = CreateObject("Microsoft.XMLDOM")
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
            "Creation of XMLDOM Failed; Err Msg : " & Err.Description
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Creation of XMLDOM Failed; Err Msg : " & Err.Description
        Call f_UnloadGlobalObjects(99)
    End If
    g_loadXml.async = False

    ' Create the XML DOM object
    Set g_loadDBXml = CreateObject("Microsoft.XMLDOM")
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
            "Creation of DB XMLDOM Failed; Err Msg : " & Err.Description
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Creation of DB XMLDOM Failed; Err Msg : " & Err.Description
        Call f_UnloadGlobalObjects(99)
    End If
    g_loadDBXml.async = False


    ' Validating the Input Parameters passed to the batch
    If Not f_ValidateInputParameter() Then
        Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
    End If
    'Read Generic Configuration Information into the dictionary object
    If Not f_ReadGenericInfo() Then
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Read 1gen"
        Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
     End If

    g_dicBatchInfo.Item("RETURN CODE") =  0
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : " & g_dicBatchInfo.Item("BatchName") & " Started"

    ' Call the function to read the configuration file
    If Not f_ReadConfigFile() Then
        Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
    End If

    ' Call the function to read the ersystcfgsystem471.xml
'''    If Not f_ReadSystemFile() Then
'''        Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
    '''End If

    ' Check if Pre-process exists
''    If g_dicBatchInfo.Item("Pre-Process") <> "" Then
        ' Call the function to read the pre-process information and execute the steps in them
''        If Not f_ExecuteProcessFile(g_arrPreBypassInfo, g_dicBatchInfo.Item("Pre-Process")) Then
''            Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
''        End If
''    End If

		 'g_dicBatchInfo.Add "FromDate",replace(DateAdd("d",-1,"21-Oct-19"),"/","")
	 
	 	 'wscript.Echo "New Run Date" & g_dicBatchInfo.Item("RunDate")

	 	 'wscript.Echo "New From Date" & g_dicBatchInfo.Item("FromDate")
		 
    ' Check if the batch name is not equal to "CMCBTCH_RUNDATE"
    If g_dicBatchInfo.Item("BatchName") <> "CMCBTCH_RUNDATE" Then
        ' Call the function to modify the batch runbook with the override parameters
        If Not f_BuildRunBook() Then
            Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
        End If
		
     End If
	 
		 
        ' Call the function to execute the batch
''        If Not f_ExecuteBatch() Then
''            Call g_dllCmnFunctions.GetLogFileName(g_connDatabase,g_dicBatchInfo.Item("ProductName"), g_dicBatchInfo.Item("ProductAppId"), g_dicBatchInfo)
''            Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
''        End If
''     End If


	 ' Check if Post-process exists
    ''If g_dicBatchInfo.Item("Post-Process") <> "" Then

        ' Call the function to read the post-process information and execute the steps in them
     ''   If Not f_ExecuteProcessFile(g_arrPostBypassInfo, g_dicBatchInfo.Item("Post-Process")) Then
       ''     Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
       '' End If

    ''End If
	

    g_dicBatchInfo.Item("RETURN CODE") =  0
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : " & g_dicBatchInfo.Item("BatchName") & " Work Runbook Generation Completed"

    ' Calling the function WriteTrace to log the completion of the batch
    ''If Not g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo) Then
      ''  f_BuildExecuteAction = False
       '' g_dicBatchInfo.Item("RETURN CODE") = 15
       '' Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
    ''End If    

    ' Disconnect the Connection to the database
   '' If Not g_dllCmnFunctions.CloseConnection(g_connDatabase) Then
     ''   Call f_UnloadGlobalObjects(-3)
    ''End If

    Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))

    ' Checking for run time errors
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
            "Err Msg : " & Err.Description
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Err Msg : " & Err.Description
        wscript.Quit (99)
    End If

End Sub

'###########################################################################
'Function Name          : f_ValidateInputParameter
'Purpose                : To check the validity of the parameters passed to
'                         main() module script
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_ValidateInputParameter()

    On Error Resume Next

    ' Local variable Declarations
    Dim l_wshShell                     ' Shell object to write the Eventlog
    Dim l_argValues                    ' Arguments object to store the command line arguments
    Dim l_nodeBatchConfig
    Dim l_nodeBatch
    Dim l_InputDate
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Log Message : Validation of Input Parameters Started"

    ' Set the f_ValidateInputParameter function to True by default
    f_ValidateInputParameter = True

    ' Create the dictionary object g_dicBatchInfo
    Set g_dicBatchInfo = CreateObject("Scripting.Dictionary")
    g_dicBatchInfo.Add "RETURN CODE", 0
    g_dicBatchInfo.Add "MESSAGE", " "
    g_dicBatchInfo.Add "SPACE", " "
    ' Creating a Shell script Object, to write the error
    Set l_wshShell = wscript.CreateObject("WScript.Shell")
	
    If Err.Number <> 0 Then
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
            "Creation of Shell Object Failed; Err Msg : " & Err.Description
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Creation of Shell Object Failed; Err Msg : " & Err.Description
        g_dicBatchInfo.Item("RETURN CODE") = 99
        f_ValidateInputParameter = False
        Exit Function
    End If

    ' Get the Command line arguments in the argument object
    Set l_argValues = wscript.Arguments


    ' Check the count of the arguments
    If l_argValues.Count = 2 Then
        ' Add the first parameter as batch name to the dictionary object
        g_dicBatchInfo.Add "BatchName", l_argValues.Item(0)

        ' Add the second parameter as database name to the dictionary object
        g_dicBatchInfo.Add "Database", l_argValues.Item(1)

        ' Add the second parameter as region id to the dictionary object
        g_dicBatchInfo.Add "RegionId", l_argValues.Item(1)

        ' Add the Third parameter as region id to the dictionary object
'        g_dicBatchInfo.Add "InputDate", l_argValues.Item(2)
 '       l_InputDate = replace(l_argValues.Item(2),"/","")
         g_dicBatchInfo.Add "RunDate", "_" + replace(CStr(Date),"/","")

    Else
        ' Setting the return code
        g_dicBatchInfo.Item("RETURN CODE") = -1

        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : Invalid No. of Parameters"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Invalid No. of Parameters"
        f_ValidateInputParameter = False
        Exit Function

    End If
wscript.Echo CStr(Date) & " -- " & wscript.ScriptFullName & Err.Number & " - " & CStr(Time()) & " - Log Message : Validation of Input Parameters crtd"
    ' Load the Config.xml in the g_loadXml object
	
wscript.Echo  g_dicBatchInfo.Item("BatchName")
    g_loadXml.Load (Mid(wscript.ScriptFullName, 1, InStrRev(wscript.ScriptFullName, "\Scripts")) + "ConfigFiles\commonsystemprop.xml")

    ' Set the node list to "BatchConfig/Category" in the variable l_batchConfig
    Set l_nodeBatchConfig = g_loadXml.documentElement.selectNodes("BatchConfig/Category")

    For Each l_nodeBatch In l_nodeBatchConfig
        ' Search for the particular batch section in the config.xml
        If l_nodeBatch.getAttributeNode("name").Value = g_dicBatchInfo.Item("BatchName") Then
            f_ValidateInputParameter = True
            Exit For
        Else
            f_ValidateInputParameter = False
        End If
    Next
	    
    ' If the batch name is not present in the Config.xml

    If Not f_ValidateInputParameter Then
        ' Setting the return code
        g_dicBatchInfo.Item("RETURN CODE") = -1
        ' Writing the Error message in the Event log
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : Invalid Batch Name"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Invalid Batch Name"
        Exit Function
    End If

    ' Destroying Command line argument object and shell object
    Set l_argValues       = Nothing
    Set l_wshShell        = Nothing
    Set l_nodeBatch       = Nothing
    Set l_nodeBatchConfig = Nothing

    ' Checking for run time errors
    If Err.Number <> 0 Then
        f_ValidateInputParameter = False
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_ValidateInputParameter Failed; Err Msg : " & Err.Description
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_ValidateInputParameter Failed; Err Msg : " & Err.Description
        g_dicBatchInfo.Item("RETURN CODE") = 99
    Else
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Log Message : Validation of Input Parameters Complete"
    End If

End Function


'###########################################################################
'Function Name          : f_ReadGenericInfo
'Purpose                : 1.To read the configuration xml, fetch all the
'                           database configuration into the dictionary
'                           objects
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_ReadGenericInfo()

    On Error Resume Next

   ' Set the f_ReadGenericInfo function to True by default
    f_ReadGenericInfo = True

    ' Declare all local variables used to access the various nodes/items in the xml
    Dim l_nodeGenConfig
    Dim l_nodeItem
    Dim l_nodeAttrib

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Log Message : Read Generic Information from Config XML Started"

    ' Set the node list to "GenericConfig/Category/Item" in the variable l_nodeGenConfig
    ' Iterate through each element in the xml and add the same to the dictionary object
    Set l_nodeGenConfig = g_loadXml.documentElement.selectNodes("GenericConfig/Category/Item")
    For Each l_nodeItem In l_nodeGenConfig
        Set l_nodeAttrib = l_nodeItem.getAttributeNode("name")
        'g_dicBatchInfo.Add l_nodeAttrib.Value, l_nodeItem.Text
        If g_dicBatchInfo.Exists(l_nodeAttrib.Value) Then
            g_dicBatchInfo.Item(l_nodeAttrib.Value) = l_nodeItem.Text
        Else
            g_dicBatchInfo.Add l_nodeAttrib.Value, l_nodeItem.Text
        End If
    Next

    ' Checking for run time errors
    If Err.Number <> 0 Then
        f_ReadGenericInfo = False
        l_wshShell.LogEvent 1, CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_ReadGenericInfo Failed; Err Msg : " & Err.Description
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_ReadGenericInfo Failed; Err Msg : " & Err.Description
        g_dicBatchInfo.Item("RETURN CODE") = 99
        Exit Function
    End If

    Set l_nodeGenConfig   = Nothing
    Set l_nodeItem        = Nothing
    Set l_nodeAttrib      = Nothing

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Log Message : Read Generic Information from Config XML Completed"

End Function

'###########################################################################
'Function Name          : f_ReadConfigFile
'Purpose                : 1.To read the configuration xml, fetch all the
'                           information in them and store them in dictionary
'                           objects and arrays
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_ReadConfigFile()

    On Error Resume Next

    ' Declare all local variables used to access the various nodes/items in the xml

    Dim l_nodeBatchConfig
    Dim l_nodeBatch
    Dim l_nodeParent
    Dim l_nodeChild
    Dim l_strConfigParam

    ' Set the f_ReadConfigFile function to True by default
    f_ReadConfigFile = True

    ' Writing the Start of the Read Config Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read Configuration XML started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " & g_dicBatchInfo.Item("MESSAGE")

    ' Create the dictionary object g_dicOverrideInfo
    Set g_dicOverrideInfo = CreateObject("Scripting.Dictionary")

    ' Set the node to batch specific information tag in the commonsystemprop.xml
    Set l_nodeBatchConfig = g_loadXml.documentElement.selectNodes("BatchConfig/Category")
    For Each l_nodeBatch In l_nodeBatchConfig
       ' Go to the particular batch section in the commonsystemprop.xml
        If l_nodeBatch.getAttributeNode("name").Value = g_dicBatchInfo.Item("BatchName") Then
           ' Fetch all the required items under the batch and store them in the g_dicBatchInfo object
            For Each l_nodeParent In l_nodeBatch.childNodes
                If l_nodeParent.nodeName = "Item" Then
                    ' If the item is not override parameters, then add the elements to the dictionary g_dicBatchInfo
                    If l_nodeParent.getAttributeNode("name").Value <> "OverrideParameters" Then
                        If l_nodeParent.hasChildnodes Then
                            l_strConfigParam = l_nodeParent.firstChild.Text
                        Else
                            l_strConfigParam = l_nodeParent.Text
                        End If
                        If Not f_ModifyParamValue(l_strConfigParam) Then
                            f_ReadConfigFile = False
                            Exit For
                        Else
                            If g_dicBatchInfo.Exists(l_nodeParent.getAttributeNode("name").Value) Then
                                g_dicBatchInfo.Item(l_nodeParent.getAttributeNode("name").Value) = l_strConfigParam
                            Else
                                g_dicBatchInfo.Add l_nodeParent.getAttributeNode("name").Value,l_strConfigParam
                            End If
                            l_strConfigParam = ""
                        End If
                    End If

                   ' If the item is pre-process/post-process, check for any bypass steps or bypass actions
                    If l_nodeParent.getAttributeNode("name").Value = "Pre-Process" Or l_nodeParent.getAttributeNode("name").Value = "Post-Process" Then
                        If Not f_BuildBypassArray(l_nodeParent, l_nodeParent.getAttributeNode("name").Value) Then
                            f_ReadConfigFile = False
                            Exit For
                        End If
                    End If

                   ' If the item is override parameters, then add the elements to the dictionary g_dicOverrideInfo
                    If l_nodeParent.getAttributeNode("name").Value = "OverrideParameters" Then
                        For Each l_nodeChild In l_nodeParent.childNodes
                            If l_nodeChild.nodeName = "Element" Then
                                l_strConfigParam = l_nodeChild.Text
                                If Not f_ModifyParamValue(l_strConfigParam) Then
                                    f_ReadConfigFile = False
                                    Exit For
                                Else
                                    g_dicOverrideInfo.Add l_nodeChild.getAttributeNode("name").Value,l_strConfigParam
                                    l_strConfigParam = ""
                                End If
                            End If
                        Next
                        If f_ReadConfigFile = False Then Exit For
                    End If
                End If
            Next
            If f_ReadConfigFile = False Then Exit For
        End If
    Next

    ' Writing the End of the Process in Database
    ' Assigning the Message item of the dictionary object
    If f_ReadConfigFile = True And Err.Number = 0 Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read Configuration XML Completed"
    Else
        f_ReadConfigFile = False
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read Configuration XML Failed"
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " & g_dicBatchInfo.Item("MESSAGE")
        ' Calling the function WriteTrace to log the Read Config Function completed message
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase, g_dicBatchInfo)
        Exit Function
    End If

    ' Destroying all the local objects
    Set l_nodeBatchConfig = Nothing
    Set l_nodeBatch = Nothing
    Set l_nodeParent = Nothing
    Set l_nodeChild = Nothing

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " & g_dicBatchInfo.Item("MESSAGE")

End Function

'###########################################################################
'Function Name          : f_ReadSystemFile
'Purpose                : 1.To read the system xml, fetch all the information
'                           in them and store them in a dictionary object
'                           g_dicBatchInfo.
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_ReadSystemFile()

    On Error Resume Next

    ' Declare local variables to access the nodes/attributes in the System default xml
    Dim l_nodeSysConfig
    Dim l_nodeParent
    Dim l_nodeAttrib

    Dim l_nodeSysConfig1
    Dim l_nodeParent1
    Dim l_nodeAttrib1
    Dim l_nodeloop
    ' Set the f_ReadSystemFile function to True by default
    f_ReadSystemFile = True

    ' Writing the Start of the Read System Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read System Default XML started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    ' Load the ersystcfgsystem.xml
     g_loadXml.Load (SYSTEM_FILE_PATH & g_dicBatchInfo.Item("RegionId") & "\ersystcfgsystem.xml")


    ' Set the node list to "Category/Item" in the variable l_nodeSysConfig
    ' Iterate through each element in the xml and add the same to the dictionary g_dicBatchInfo
    Set l_nodeSysConfig = g_loadXml.documentElement.selectNodes("Category/Item")
    For Each l_nodeParent In l_nodeSysConfig
        Set l_nodeAttrib = l_nodeParent.getAttributeNode("name")
        If Not g_dicBatchInfo.Exists(l_nodeAttrib.Value) Then
            g_dicBatchInfo.Add l_nodeAttrib.Value, l_nodeParent.Text
        End If
    Next

    ' Writing the End of the Process in Database
    ' Assigning the Message item of the dictionary object
    If Err.Number = 0 Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read System Default XML Completed"
    Else
        f_ReadSystemFile = False
	'msgbox Err.Description
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read System Default XML Failed"
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        ' Calling the function WriteTrace to log the Read System Function completed message
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
        Exit Function
    End If

    ' Destroying all the local objects
    Set l_nodeSysConfig = Nothing
    Set l_nodeParent = Nothing
    Set l_nodeAttrib = Nothing

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

End Function

'###########################################################################
'Function Name          : f_ReadSystemDBFile
'Purpose                : 1.To read the system xml, fetch all the DB information
'                           in them and store them in a dictionary object
'                           g_dicBatchInfo.
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_ReadSystemDBFile()

    On Error Resume Next
    ' Declare local variables to access the nodes/attributes in the System default xml
    Dim l_nodeSys_Config
    Dim l_node_Parent
    Dim l_node_Attrib


    ' Set the f_ReadSystemFile function to True by default
    f_ReadSystemDBFile = True

    ' Writing the Start of the Read System DB Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read Database Details from System Default XML started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    ' Load the ersystcfgsystem.xml
     g_loadDBXml.Load (SYSTEM_FILE_PATH & g_dicBatchInfo.Item("RegionId") & "\ersystcfgsystem.xml")

    ' Set the node list to "Category/DataDomain/ConnectionSpecification/ConnectionAttribute" in the variable l_nodeSys_Config

      Set l_nodeSys_Config = g_loadDBXml.documentElement.selectNodes("Category/DataDomain/ConnectionSpecification/ConnectionAttribute")
    	For Each l_node_Parent In l_nodeSys_Config
       	Set l_node_Attrib = l_node_Parent.getAttributeNode("name")
        If Not g_dicBatchInfo.Exists("SYS"& l_node_Attrib.Value) Then
            g_dicBatchInfo.Add "SYS"& l_node_Attrib.Value, l_node_Parent.Text
        End If
   Next

    ' Writing the End of the Process in Database
    ' Assigning the Message item of the dictionary object
    If Err.Number = 0 Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read Database Details from System Default XML Completed"
    Else
        f_ReadSystemDBFile = False    
				'msgbox Err.Description
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Read Database Details from System Default XML Failed"
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        ' Calling the function WriteTrace to log the Read System Function completed message
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
        Exit Function
    End If

    ' Destroying all the local objects
    Set l_nodeSys_Config = Nothing
    Set l_node_Parent = Nothing
    Set l_node_Attrib = Nothing

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

End Function

'###########################################################################
'Function Name          : f_ExecuteProcessFile
'Purpose                : 1.To read the various steps in the pre/post process
'                           xml and executes the steps in them one by one if
'                           the step is not present in the g_arrPreBypassInfo/
'                           g_arrPostBypassInfo array
'Input                  : (ByVal)p_bypassArray - pre/post bypass array
'                         (ByVal)p_processFileName - pre/post process file name
'Output                 : True/False
'###########################################################################
Private Function f_ExecuteProcessFile(ByVal p_bypassArray, ByVal p_processFileName)

    On Error Resume Next

    ' Local variable Declarations
    Dim l_nodeStepList
    Dim l_nodeStepNumber
    Dim l_blnChkStep
    Dim l_blnChkAction
    Dim l_nodeProc
    Dim l_nodeStep
    Dim l_nodeAction
    Dim l_nodeActChild
    Dim l_nodeActNumber
    Dim l_nodeActInfo

    ' Set the f_ExecuteProcessFile function to True by default
    f_ExecuteProcessFile = True

    ' Writing the Start of the Execute Pre/Post Process Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute Pre/Post Process " & p_processFileName & " Started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    ' Load the pre/post process xml file
    g_loadXml.Load (g_dicBatchInfo.Item("ConfigFilesPath") + "\" + p_processFileName)

    ' Set the node to root node
    Set l_nodeProc = g_loadXml.documentElement

    ' Get all the steps into a local variable l_stepList
    Set l_nodeStepList = l_nodeProc.getElementsByTagName("Step")
    ' For each step in the list of all steps
    For Each l_nodeStep In l_nodeStepList
        ' Get the step number
        Set l_nodeStepNumber = l_nodeStep.getAttributeNode("number")

        ' Check if the step number is present in the array
        l_blnChkStep = f_CheckBypass(l_nodeStepNumber.Value, p_bypassArray)

        ' If the step number is not present in the bypass array
        If Not (l_blnChkStep) Then
            ' Set the node to Action to fetch all the action numbers in a step
            Set l_nodeAction = l_nodeStep.getElementsByTagName("Action")

            For Each l_nodeActChild In l_nodeAction
                ' Get the action number
                Set l_nodeActNumber = l_nodeActChild.getAttributeNode("number")

                ' Check if the bypass action is present in the array
                l_blnChkAction = f_CheckBypass(l_nodeActNumber.Value, p_bypassArray)

                ' If the action number is not present in the bypass array
                If Not (l_blnChkAction) Then
                    ' Fetch the function/vbscript/procedure name and parameters from the pre/post proc xml - start
                    For Each l_nodeActInfo In l_nodeActChild.childNodes
                        ' If the pre/post process step is a DLL function
                        ' Fetch the function name and parameters
                        If l_nodeActInfo.nodeName = "FunctionInfo" Or l_nodeActInfo.nodeName = "ScriptInfo" Or l_nodeActInfo.nodeName = "SqlInfo" Then
                            wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  "Log Message : Start of Step/Action No: " & l_nodeActNumber.Value & " in the Pre/Post Process File " & p_processFileName
                            If Not f_BuildExecuteAction(l_nodeActInfo, CStr(l_nodeActInfo.nodeName)) Then
                                f_ExecuteProcessFile = False
                                wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  "Log Message : Failed in Step/Action No: " & l_nodeActNumber.Value & " in the Pre/Post Process File " & p_processFileName
                                Exit For
                            End If
                            wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  "Log Message : End of Step/Action No: " & l_nodeActNumber.Value & " in the Pre/Post Process File " & p_processFileName
                        End If
                    Next
                    If f_ExecuteProcessFile = False Then Exit For
                Else
                    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  "Log Message : Bypassed Action No: " & l_nodeActNumber.Value & " in the Pre/Post Process File " & p_processFileName
                End If
            Next
            If f_ExecuteProcessFile = False Then Exit For
        Else
            wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  "Log Message : Bypassed Step No: " & l_nodeStepNumber.Value & " in the Pre/Post Process File " & p_processFileName
        End If
    Next

    ' Writing the End of the Execute Pre/Post Process Function in Database
    ' Assigning the Message item of the dictionary object
    If f_ExecuteProcessFile = True Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute Pre/Post Process " & p_processFileName & " Completed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
    Else
        f_ExecuteProcessFile = False
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute Pre/Post Process " & p_processFileName & " Failed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        ' Calling the function WriteTrace to log the Execute Pre/Post Process Function completed message
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
        Exit Function
    End If

    ' Destroying all the local objects
    Set l_nodeStepList = Nothing
    Set l_nodeStepNumber = Nothing
    Set l_nodeProc = Nothing
    Set l_nodeStep = Nothing
    Set l_nodeAction = Nothing
    Set l_nodeActChild = Nothing
    Set l_nodeActNumber = Nothing
    Set l_nodeActInfo = Nothing

    ' Checking for run time errors
    If Err.Number <> 0 Then
        f_ExecuteProcessFile = False
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_ExecuteProcessFile Failed; Err Msg : " & Err.Description
    End If

End Function

'###########################################################################
'Function Name          : f_BuildRunBook
'Purpose                : 1.To modify the runbook with the override parameters
'                           present in the g_dicOverrideInfo dictionary object
'                           and generic parameters present in the g_dicBatchInfo
'                           dictionary object and save the modified runbook with
'                           the name <database>.runbook name in a location
'                           configured in the configuration fileTo read the
'                           various steps in the pre/post process
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_BuildRunBook()

    On Error Resume Next

    ' Local variable Declarations
    Dim l_strRunBook
    Dim l_strRunBookLoc
    Dim l_nodeRoot
    Dim l_nodeItems
    Dim l_nodeItem
    Dim l_nodeOldText
    Dim l_nodeItemAttrib
    Dim l_nodeNewText



    ' Set the f_BuildRunBook function to True by default
    f_BuildRunBook = True

    ' Writing the Start of the Build Runbook Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Runbook Process started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    ' Fetch the run book name from the dictionary object
    l_strRunBook = g_dicOverrideInfo.Item("RunBookName")

    ' Fetch the location of the run book from the dictionary object
    l_strRunBookLoc = g_dicBatchInfo.Item("RunBookPath")

    ' Load the run book xml into the DOM object g_loadXml
    g_loadXml.Load (l_strRunBookLoc + "\" + l_strRunBook)

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  l_strRunBook & "   " & l_strRunBookLoc

    ' Set the node to root node in the runbook
    Set l_nodeRoot = g_loadXml.documentElement

    ' Fetch all the items under the root node
    Set l_nodeItems = l_nodeRoot.getElementsByTagName("Item")

    ' Fetch the text present in the runbook xml for each item in l_nodeItems

    For Each l_nodeItem In l_nodeItems
        Set l_nodeOldText = l_nodeItem.firstChild
        Set l_nodeItemAttrib = l_nodeItem.getAttributeNode("name")
        ' Check the value of the item in the override info dictionary object
        If (g_dicOverrideInfo.Item(l_nodeItemAttrib.Value) <> "") Then
            Set l_nodeNewText = g_loadXml.createTextNode(g_dicOverrideInfo.Item(l_nodeItemAttrib.Value))
            If l_nodeItem.hasChildnodes Then
                l_nodeItem.replaceChild l_nodeNewText, l_nodeOldText
            Else
                l_nodeItem.AppendChild l_nodeNewText
            End If
        ' Check the value of the item in the batch info dictionary object
        ElseIf (g_dicBatchInfo.Item(l_nodeItemAttrib.Value) <> "") Then
            Set l_nodeNewText = g_loadXml.createTextNode(g_dicBatchInfo.Item(l_nodeItemAttrib.Value))
            If l_nodeItem.hasChildnodes Then
                l_nodeItem.replaceChild l_nodeNewText, l_nodeOldText
            Else
                l_nodeItem.AppendChild l_nodeNewText
            End If
        End If
    Next

'    l_strRunBook = Mid(g_dicBatchInfo.Item("Database"),InStr(1,g_dicBatchInfo.Item("Database"),"facets") + 6) + "." + l_strRunBook
    l_strRunBook = g_dicBatchInfo.Item("Database") + "." + l_strRunBook

    ' Save the modified run book xml
    g_loadXml.save (g_dicBatchInfo.Item("WorkfilePath") + "\" + l_strRunBook)
    'g_dicOverrideInfo.Item("RunBookName") = g_dicBatchInfo.Item("WorkfilePath") + "\" + l_strRunBook

    ' Writing the End of the Build Runbook Function in Database
    ' Assigning the Message item of the dictionary object
    If Err.Number = 0 Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Runbook Process Completed"
    Else
        f_BuildRunBook = False
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Runbook Process Failed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        g_dicBatchInfo.Item("RETURN CODE") = 99
        ' Calling the function WriteTrace to log the Build Runbook Function completed message
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
        Exit Function
    End If

    ' Destroying all the local objects
    Set l_nodeRoot = Nothing
    Set l_nodeItems = Nothing
    Set l_nodeItem = Nothing
    Set l_nodeOldText = Nothing
    Set l_nodeItemAttrib = Nothing
    Set l_nodeNewText = Nothing

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

End Function

'###########################################################################
'Function Name          : f_ExecuteBatch
'Purpose                : 1.To execute the batch with the modified run book xml
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_ExecuteBatch()

    On Error Resume Next

    ' Local variable Declarations
    Dim l_wshShell                     ' Shell object to execute the batch
    Dim l_strRunBook
    Dim l_strRunBookLoc
    Dim l_intReturnCode
    Dim l_strModRunBook

    ' Set the function to True by default
    f_ExecuteBatch = True



    ' Writing the Start of the Execute Batch Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute Batch Process started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    ' Creating a Shell script Object
    Set l_wshShell = wscript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        g_dicBatchInfo.Item("RETURN CODE") = 99
        f_ExecuteBatch = False
        Exit Function
    End If

    ' Fetch the run book name from the dictionary object
    l_strRunBook = g_dicOverrideInfo.Item("RunBookName")

    ' Fetch the location of the run book from the dictionary object
    l_strRunBookLoc = g_dicBatchInfo.Item("WorkfilePath")

    ' Form the modified run book name along with path
    l_strModRunBook = l_strRunBookLoc + "\" + g_dicBatchInfo.Item("Database") + "." + l_strRunBook

    ' Run the batch using the shell object
    l_intReturnCode = l_wshShell.Run("C:\Windows\SysWow64\cscript.exe " & g_dicBatchInfo("ProcFileDirectory") &  "\ErSys0FrmExecuteJob.wsf " & "--RUNBOOK=" & l_strModRunBook & " --Region=" & g_dicBatchInfo.Item("RegionId") , 1, True)

    ' Assign the return code as zero for BOTB batch - 05062019
	
	If l_intReturnCode = 8 and g_dicBatchInfo.Item("BatchName") = "MultiplanHIPAABotbBatch" Then
        l_intReturnCode = 0
    End If

		
	If l_intReturnCode = 8 and g_dicBatchInfo.Item("BatchName") = "FIRSTSOURCEBOTBBatch" Then
        l_intReturnCode = 0
    End If
    
    ' Assign the return code as zero for Pricer BOTB batch - 09122019
    	If l_intReturnCode = 8 and g_dicBatchInfo.Item("BatchName") = "RepricerBotbBatch" Then
        l_intReturnCode = 0
    End If
	
    ' Writing the End of the Execute Batch Function in Database
    ' Assigning the Message item of the dictionary object

    If Not (g_dicBatchInfo.Exists("ThresholdReturnCode")) Then
        g_dicBatchInfo.Item("ThresholdReturnCode") = 0
    End If

    'If l_intReturnCode = 0 or l_intReturnCode < CInt(g_dicBatchInfo.Item("ThresholdReturnCode")) Then   ' Commented for PM10362613 on 06/13/2016
    If l_intReturnCode = 0 Then  ' Added for PM10362613 on 06/13/2016
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute Batch Process Completed"
    Else
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute Batch Process Failed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        g_dicBatchInfo.Item("RETURN CODE") = l_intReturnCode
        f_ExecuteBatch = False
        ' Calling the function WriteTrace to log the Execute Batch Function completed message
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
        Exit Function
    End If

    ' Destroy the shell object
    Set l_wshShell = Nothing

    ' Checking for run time errors
    If Err.Number <> 0 Then
        f_ExecuteBatch = False
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_ExecuteBatch Failed; Err Msg : " & Err.Description
    Else
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
    End If

End Function

'###########################################################################
'Function Name          : f_Execute837Batch
'Purpose                : 1.To execute the HIPAA Inbound EDI 837 batch for
'                           a set of input files
'Input                  : None
'Output                 : True/False
'###########################################################################
Private Function f_Execute837Batch()

    On Error Resume Next

    ' Local variable Declarations
    Dim l_strWorkFilePath
    Dim l_strOutputDir
    Dim l_fsoInputFileList
    Dim l_strReadFile
    Dim l_strFileContents
    Dim l_intColPos
    Dim l_strInputFile

    ' Set the function to True by default
    f_Execute837Batch = True

    ' Writing the Start of the Execute 837 Batch Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute 837 Batch Process started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    ' Fetch the work file path from the g_dicBatchInfo object
    l_strWorkFilePath = g_dicBatchInfo.Item("WorkfilePath")

    ' Fetch the output directory from the g_dicOverrideInfo object
    l_strOutputDir = g_dicOverrideInfo.Item("OutputDir")

    ' Create the File System Object and open the input file list
    Set l_fsoInputFileList = CreateObject("Scripting.FileSystemObject")
    Set l_strReadFile = l_fsoInputFileList.OpenTextFile(l_strWorkFilePath + "\" + g_dicBatchInfo.Item("Database") + ".837InputFileList.txt", 1, False)

    ' Read each line and perform the following for each input file
    '12/15/2016 PRB0125349 - commented the do while loop
    ' Do While Not(l_strReadFile.AtEndOfStream)
    '    l_strFileContents = l_strReadFile.ReadLine
    '    l_intColPos = InStr(1, l_strFileContents, g_dicOverrideInfo.Item("InputFileExtension"), 1)
        ' Fetch the input file name without extension
    '    l_strInputFile = Left(l_strFileContents, l_intColPos - 2)
        ' Update the override parameter "InputFile" present in the g_dicOverrideInfo object with the file name without the extension
    '    g_dicOverrideInfo.Item("InputFile") = l_strInputFile

        ' Call the f_BuildRunBook function
        If Not f_BuildRunBook() Then
            f_Execute837Batch = False
            g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Run Book Failed"
            wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
            ' Exit Do 12/15/2016 PRB0125349 - commented the do while loop
        End If

        ' Execute the batch with the modified run book
        If Not f_ExecuteBatch() Then
            f_Execute837Batch = False
            g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute 837 Batch Failed"
            wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
            ' Exit Do 12/15/2016 PRB0125349 - commented the do while loop
        Else
            Call g_dllCmnFunctions.GetLogFileName(g_connDatabase,g_dicBatchInfo.Item("ProductName"), g_dicBatchInfo.Item("ProductAppId"), g_dicBatchInfo)
        End If
   ' Loop

    '-----------------------------------------------------------------------
    ' Writing the End of the Execute 837 Batch Function in Database
    '-----------------------------------------------------------------------
    ' Assigning the Message item of the dictionary object
    If f_Execute837Batch = True Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Execute 837 Batch Process completed"
    Else
        f_Execute837Batch = False
        g_dicBatchInfo.Item("RETURN CODE") = 99
        ' Calling the function WriteTrace to log the Execute 837 Batch Function completed message
        'Call g_dllCmnFunctions.WriteTrace(g_connDatabase, g_dicBatchInfo)
        Exit Function
    End If

    ' Destroy the local objects
    Set l_fsoInputFileList = Nothing

    ' Checking for run time errors
    If Err.Number <> 0 Then
        f_Execute837Batch = False
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_Execute837Batch Failed; Err Msg : " & Err.Description
    Else
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
    End If

End Function

'###########################################################################
'Function Name          : f_ModifyParamValue
'Purpose                : 1.To update the Parameter Value
'Input                  : (ByRef) p_strParamValue
'Output                 : True/False
'###########################################################################
Private Function f_ModifyParamValue(ByRef p_strParamValue)

    On Error Resume Next

    ' Local variable Declarations
    Dim l_strFuncParam      ' Variable to store the parameter value of a function
    Dim l_intPlusPos        ' Variable to store the position of first occurrence of "+" in the parameter
    Dim l_arrParameter      ' Array to store the different strings separated by "+" present in a parameter
    Dim l_intStart          ' Variable to store the start position of the brace
    Dim l_intEnd            ' Variable to store the end position of the brace
    Dim l_Item              ' Variable to store item name
    Dim l_intNoOfParam      ' Variable to store the count of input parameter
    Dim l_intArrayIndex
    Dim l_intElementIndex
    Dim l_strModParamValue

    ' Set the function to True by default
    f_ModifyParamValue = True

    ' Writing the Start of the Modify Param Value Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Process started for modifying the value of the parameter: " + p_strParamValue
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    ' To check if the string "+" is present in the parameter, get the position of the string "+"
    l_intPlusPos = InStr(1, p_strParamValue, "+", 1)
    If l_intPlusPos > 0 Then
        ' Store the strings present in the parameter into the array
        l_arrParameter = Split(p_strParamValue, "+")
    Else
        ReDim l_arrParameter(0)
        l_arrParameter(0) = p_strParamValue
    End If

    ' Loop to fetch parameters from dictionary objects
    For l_intElementIndex = 0 To UBound(l_arrParameter)
        If (InStr(1, l_arrParameter(l_intElementIndex), "Config(", 1) > 0) Then
            l_intStart = InStr(1, l_arrParameter(l_intElementIndex), "(", 1)
            l_intEnd = InStr(1, l_arrParameter(l_intElementIndex), ")", 1)
            l_Item = Mid(l_arrParameter(l_intElementIndex), l_intStart + 2, l_intEnd - l_intStart - 3)
            If (g_dicBatchInfo.Item(l_Item) <> "") Then
                l_arrParameter(l_intElementIndex) = g_dicBatchInfo.Item(l_Item)
            ElseIf (g_dicOverrideInfo.Item(l_Item) <> "") Then
                l_arrParameter(l_intElementIndex) = g_dicOverrideInfo.Item(l_Item)
            Else
                l_arrParameter(l_intElementIndex) = ""
            End If
        ElseIf (InStr(1, l_arrParameter(l_intElementIndex), "mmddyy.hhmmss", 1) > 0) Then
            l_arrParameter(l_intElementIndex) = CStr(DatePart("m",Now)) + CStr(DatePart("d",Now)) + CStr(DatePart("yyyy",Now)) + "." + CStr(DatePart("h",Now)) + CStr(DatePart("n",Now)) + CStr(DatePart("s",Now))
        ElseIf (InStr(1, l_arrParameter(l_intElementIndex), "mm.dd.yy.Thh", 1) > 0) Then
            l_arrParameter(l_intElementIndex) = CStr(DatePart("m",Now)) + "." + CStr(DatePart("d",Now)) + "." + CStr(DatePart("yyyy",Now)) + ".T" + CStr(DatePart("h",Now))
        End If
        
'        If(Trim(l_arrParameter(l_intElementIndex)) = g_dicBatchInfo.Item("InputDate") ) Then
'         l_strModParamValue = l_strModParamValue + Trim(l_arrParameter(l_intElementIndex)) + " "
'        Else
        'Bulding of parameter by concatenating each element in the array
        l_strModParamValue = Trim(l_strModParamValue) + Trim(l_arrParameter(l_intElementIndex))
'        End If
    Next
'        l_strModParamValue = Trim(l_strModParamValue)
    '-----------------------------------------------------------------------
    ' Writing the End of the ModifyParamValue Function in Database
    '-----------------------------------------------------------------------
    ' Assigning the Message item of the dictionary object
    If Err.Number = 0 Then
        If p_strParamValue <> l_strModParamValue Then
            g_dicBatchInfo.Item("MESSAGE") = "Log Message : Successfully modified the value for the parameter " & p_strParamValue & " to " & l_strModParamValue
        Else
            g_dicBatchInfo.Item("MESSAGE") = "Log Message : No change in the value for the parameter " & p_strParamValue
        End If
    Else
        f_ModifyParamValue = False
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Failed to modify the value of the parameter " & p_strParamValue
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        ' Calling the function WriteTrace to log the Modify Param Value
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
        Exit Function
    End If

    'Update the input parameter with the modifed value
    p_strParamValue = l_strModParamValue

    ' Destroy the local objects
    Set l_arrParameter = Nothing

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

End Function

'###########################################################################
'Function Name          : f_BuildBypassArray
'Purpose                : 1.To build the pre/post process bypass array
'                           which contains all the bypass steps or actions
'                           for the pre/post processes.
'Input                  : (ByVal) p_strNode
'                         (ByVal) p_strProcessType
'Output                 : True/False
'###########################################################################
Private Function f_BuildBypassArray(ByVal p_strNode, ByVal p_strProcessType)

    On Error Resume Next

    ' Local variable Declarations
    Dim l_nodeBypassItems
    Dim l_intNoOfItems
    Dim l_intArrayLen       ' Variable to store the array dimension
    Dim l_arrBypassInfo
    Dim l_nodeChild

    ' Set the function to True by default
    f_BuildBypassArray = True

    ' Writing the Start of the Build Bypass Array Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Bypass Array Process started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    Set l_nodeBypassItems = p_strNode.getElementsByTagName("Element")
    l_intNoOfItems = l_nodeBypassItems.length
    ' If there are any bypass items
    If p_strNode.hasChildnodes And l_intNoOfItems > 1 Then
    ' Set the array dimension to be equal to the no of bypass items
        ReDim l_arrBypassInfo(l_intNoOfItems - 2)
        l_intArrayLen = 0       ' Initialize the variable to zero
        For Each l_nodeChild In p_strNode.childNodes
        ' Fetch each bypass item and store it in the array
            If (l_nodeChild.nodeName = "Element") Then
                If (l_nodeChild.getAttributeNode("name").Value = "BypassStep" Or l_nodeChild.getAttributeNode("name").Value = "BypassAction") Then
                    l_arrBypassInfo(l_intArrayLen) = l_nodeChild.Text
                    l_intArrayLen = l_intArrayLen + 1
                End If
            End If
        Next
    Else
        ' Redimension to zero
        ReDim l_arrBypassInfo(0)
    End If

    If p_strProcessType = "Pre-Process" Then
        g_arrPreBypassInfo = l_arrBypassInfo
    Else
        g_arrPostBypassInfo = l_arrBypassInfo
    End If

    ' Writing the End of the Build Bypass Array Function in Database
    ' Assigning the Message item of the dictionary object
    If Err.Number = 0 Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Bypass Array Process completed"
    Else
        f_BuildBypassArray = False
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Bypass Array Process Failed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        g_dicBatchInfo.Item("RETURN CODE") = 99
        ' Calling the function WriteTrace to log the Build Bypass Array Function completed message
        Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
        Exit Function
    End If

    ' Destroy the local objects
    Set l_nodeBypassItems = Nothing
    Set l_arrBypassInfo = Nothing
    Set l_nodeChild = Nothing

    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

End Function

'###########################################################################
'Function Name          : f_CheckBypass
'Purpose                : 1.To check if the bypass step or action is present
'                           in the pre/post bypass array
'Input                  : (ByVal) p_bypassStep
'                         (ByVal) p_bypassArray
'Output                 : True/False
'###########################################################################
Private Function f_CheckBypass(ByVal p_bypassStep, ByVal p_bypassArray)

    On Error Resume Next

    ' Declare local variables
    Dim l_intArrayIndex

    ' Set the function to false by default
    f_CheckBypass = False

    ' Check if the step number is present in the array
    For l_intArrayIndex = 0 To UBound(p_bypassArray)
        If p_bypassStep = p_bypassArray(l_intArrayIndex) Then
            f_CheckBypass = True
            Exit For
        End If
    Next

End Function

'###########################################################################
'Function Name          : f_BuildExecuteAction
'Purpose                : 1.To the VB DLL function or VB script or stored
'                           procedure present as step/action in the pre/post
'                           process xmls
'Input                  : (ByVal) p_strActInfo
'                         (ByVal) p_strType
'Output                 : True/False
'###########################################################################
Private Function f_BuildExecuteAction(ByVal p_strActInfo, ByVal p_strType)

    On Error Resume Next

    ' Declare local variables
    Dim l_wshShell
    Dim l_strName           'Variable to store the function/vb script/procedure name
    Dim l_intNoOfParam
    Dim l_nodeLastChild
    Dim l_nodeParam
    Dim l_strScriptParam    'String to store the paramters of a vb script
    Dim l_arrParamList      'Array to store the parameters of a function/vb script
    Dim l_intArrayIndexOne  'Variables to iterate through the two-dimensionsal array
    Dim l_intArrayIndex
    Dim l_intReturnCode
    Dim l_strArrayParam
    Dim l_dicFuncParam
    Dim l_strParamValue
    Dim l_rstResult

'    Const l_conadVarChar = 200              ' Constant variable to store the value equivalent to varchar
    Const l_conadVarChar =3
    Const l_conadParamInput = 1             ' Constant variable to store the value equivalent to Input paramter for the stored procedure

    ' Set the function to True by default
    f_BuildExecuteAction = True
    l_intReturnCode = 0


    ' Creating a Shell script Object, to write the error
    Set l_wshShell = wscript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        g_dicBatchInfo.Item("RETURN CODE") = 99
        f_BuildExecuteAction = False
        Exit Function
    End If

    ' Check if p_strActInfo has any child nodes
    If p_strActInfo.hasChildnodes() Then
        l_strName = CStr(p_strActInfo.firstChild.Text)
        If p_strActInfo.firstChild.Text = p_strActInfo.lastChild.Text Then
            l_intNoOfParam = 0
            l_intArrayIndexOne = -1
        Else
            Set l_nodeLastChild = p_strActInfo.lastChild
            ' Get the number of parameters for the function
            l_intNoOfParam = l_nodeLastChild.childNodes.length
        End If
    Else
        l_intNoOfParam = 0
        l_intArrayIndexOne = -1
    End If


    ' Writing the Start of the Build Execute Action Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Input Parameter List for the " & p_strType & "-" & l_strName & " started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

   ' If number of parameters is greater than zero
    If l_intNoOfParam > 0 Then
        ' Form an array for storing the parameters for the function/vb script/procedure
        Set l_dicFuncParam = CreateObject("Scripting.Dictionary")
        ' Initialize the array index variables to zero
        ReDim l_arrParamList(l_intNoOfParam, 2)
        l_intArrayIndexOne = 0
        ' Populate the dictionary object with the parameters of the function/vb script
        For Each l_nodeParam In l_nodeLastChild.childNodes
            l_strParamValue = CStr(l_nodeParam.lastChild.Text)
            If f_ModifyParamValue(l_strParamValue) Then
            Select Case p_strType
            Case "FunctionInfo"
                l_dicFuncParam.Add CStr(l_nodeParam.firstChild.Text),l_strParamValue
            Case "ScriptInfo"
                l_strArrayParam = l_strArrayParam + " " + l_strParamValue
            Case "SqlInfo"
                    l_arrParamList(l_intArrayIndexOne, 0) = CStr(l_nodeParam.firstChild.Text)
                    l_arrParamList(l_intArrayIndexOne, 1) = l_conadVarChar & "," & l_conadParamInput & "," & Len(Trim(l_strParamValue)) & "," & l_strParamValue
                    l_intArrayIndexOne = l_intArrayIndexOne + 1
                End Select
            Else
                f_BuildExecuteAction = False
            End If
        Next
    Else
        ReDim l_arrParamList(0,2)
    End If


    If f_BuildExecuteAction = True Then
        ' Writing the Start of the Build Execute Action Function in Database
        ' Assigning the Message item of the dictionary object
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Input Parameter List for the " & p_strType & "-" & l_strName & " completed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

        ' Execution of VB DLL Function
        If p_strType = "FunctionInfo" Then
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Input Parameter List for the " & p_strType & "-" & l_strName & " completed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

            If Not g_dllCmnFunctions.CallDLLFunction(l_strName,l_dicFuncParam,g_connDatabase,g_dicBatchInfo,g_dicOverrideInfo) Then
                If l_strName = "CheckInputFile" and  g_dicBatchInfo.Item("RETURN CODE") = 3 Then
                    f_BuildExecuteAction = True
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & g_dicBatchInfo.Item("MESSAGE")
                    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                    If g_dicBatchInfo.Item("BatchName") = "JDHC_CCSUMI0_DLYCORR_RESUB" or g_dicBatchInfo.Item("BatchName") = "JDHC_CCSUMI0_DLYCORR_CONEBHRESUB" or g_dicBatchInfo.Item("BatchName") = "JDHC_CMCXC00_CLAIMS_LOAD" Then
                        g_dicBatchInfo.Item("RETURN CODE") = 0
                    End If
            ' 11/20/2009 - Cognizant Offshore Modified - begin
                ' If there are no input files available for mschip MMS quit the batch without any error message
                    If g_dicBatchInfo.Item("BatchName") = "JDHC_CCSMMS0_MSCHIP" Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "MSCHIP - No Input Files are available for processing: "
                        g_dicBatchInfo.Item("RETURN CODE") = 0
                    End If
            ' 11/20/2009 - Cognizant Offshore Modified - end

            ' 09/14/2011 - Cognizant Offshore Modified - begin
                ' If there are no input files available for Care One Medical Outpatient quit the batch without any error message
                    If g_dicBatchInfo.Item("BatchName") = "JDHC_CCSUMI0_DLYNORM_CONEMDRF" Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "Care one Medical Outpatient Batch - No Input Files are available for processing:"
                        g_dicBatchInfo.Item("RETURN CODE") = 0
                    End If
            ' 09/14/2011 - Cognizant Offshore Modified - end
            ' 09/14/2011 - Cognizant Offshore Modified - begin
                ' If there are no input files available for Care Planner Outpatient Referral quit the batch without any error message
                    If g_dicBatchInfo.Item("BatchName") = "JDHC_CCSUMI0_DLYNORM_CAREREFF" Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "CarePlanner UM batch Outpatient Referral - No Input Files are available for processing: "
                        g_dicBatchInfo.Item("RETURN CODE") = 0
                    End If
            ' 09/14/2011 - Cognizant Offshore Modified - end
            ' 09/14/2011 - Cognizant Offshore Modified - begin
                ' If there are no input files available for Careplanner UM Daily Batch quit the batch without any error message
                    If g_dicBatchInfo.Item("BatchName") = "JDHC_CCSUMI0_DLYNORM_CARE" Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "Careplanner UM Daily bacth - No Input Files are available for processing: "
                        g_dicBatchInfo.Item("RETURN CODE") = 0
                    End If
            ' 09/14/2011 - Cognizant Offshore Modified - end
            ' 09/14/2011 - Cognizant Offshore Modified - begin
                ' If there are no input files available for Careone Medical Inpatient UM Daily batch quit the batch without any error message
                    If g_dicBatchInfo.Item("BatchName") = "JDHC_CCSUMI0_DLYNORM_CONEMDPR" Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "Careone Medical Inpatient UM Daily batch - No Input Files are available for processing: "
                        g_dicBatchInfo.Item("RETURN CODE") = 0
                    End If
            ' 09/14/2011 - Cognizant Offshore Modified - end


                    Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
        ' 04/21/2006 - Cognizant Offshore - Modified
                ElseIf l_strName = "UpdateDicOverride" and (g_dicBatchInfo.Item("BatchName") = "JDHC_CMCBIL0" or g_dicBatchInfo.Item("BatchName") = "JDHC_CDSBIL0" or g_dicBatchInfo.Item("BatchName") = "JDHC_CDSFBA0") and g_dicBatchInfo.Item("RETURN CODE") = 17 Then
                    f_BuildExecuteAction = True
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & g_dicBatchInfo.Item("MESSAGE")
                    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                    g_dicBatchInfo.Item("RETURN CODE") = 0
                    Call f_UnloadGlobalObjects(0)
                Else
                    f_BuildExecuteAction = False
                    g_dicBatchInfo.Item("RETURN CODE") = 15
                    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Failed to execute the function " & l_strName
                    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                End If
            End If
        ' Execution of vb script
        ElseIf p_strType = "ScriptInfo" Then
            ' Execute the vb script using shell function passing the parameters
            'l_intReturnCode = l_wshShell.Run("CScript " & g_dicBatchInfo.Item("ScriptsPath") & "\" & l_strName & " " & l_strArrayParam, 1, True)
            l_intReturnCode = l_wshShell.Run("C:\Windows\SysWow64\cscript.exe " & g_dicBatchInfo.Item("ScriptsPath") & "\" & l_strName & " " & l_strArrayParam, 1, True)

            If l_intReturnCode <> 0 and l_strName <> "FA_CopyDelArch_SourceToDest.vbs" Then
                f_BuildExecuteAction = False
                g_dicBatchInfo.Item("RETURN CODE") = 16
                g_dicBatchInfo.Item("MESSAGE") = "Log Message : Failed to execute the Script " & l_strName
                wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
            ' If there are no input files available for 837I quit the batch without any error message
            ElseIf l_intReturnCode <> 0 and l_strName = "FA_CopyDelArch_SourceToDest.vbs" Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "eConn - No Input Files are available for processing: "
                    g_dicBatchInfo.Item("MESSAGE") = "Log Message : No Input Files are available for processing"
                    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                    g_dicBatchInfo.Item("RETURN CODE") = 0
                    'Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
              
                   ' wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                    l_intReturnCode = l_wshShell.Run("C:\Windows\SysWow64\cscript.exe " & g_dicBatchInfo.Item("ScriptsPath") & "\FA_NTBTCHFRMWK_SendEmail.vbs " & g_dicBatchInfo.Item("Mail Path") & " NOTIFICATION_837I.txt " & g_dicBatchInfo.Item("Mail_to_ok") & " " & g_dicBatchInfo.Item("Sender") & " """ & g_dicBatchInfo.Item("Subject") & """ " & g_dicBatchInfo.Item("Mail_ServerName") & " " & g_dicBatchInfo.Item("Mail_Port"), 1, True)
                    If l_intReturnCode = 0 and l_strName = "FA_CopyDelArch_SourceToDest.vbs" Then
                       g_dicBatchInfo.Item("MESSAGE") = "Log Message : Notification Email sent successfully!"
                       wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                    End If
            Call f_UnloadGlobalObjects(0)
            ElseIf l_strName = "JS_BuildInputFile.vbs" Then
                If Not g_dllCmnFunctions.CheckInputFile(g_dicBatchInfo.Item("Database") + ".ecmms.inputfiles",g_dicBatchInfo.Item("WorkfilePath"),g_connDatabase,g_dicBatchInfo) Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "eConn - No Input Files are available for processing: "
                    g_dicBatchInfo.Item("MESSAGE") = "Log Message : No Input Files are available for processing"
                    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                    g_dicBatchInfo.Item("RETURN CODE") = 0
                    'Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
                    Call f_UnloadGlobalObjects(0)
                End If
                ' 07/08/2010 - Cognizant Offshore Modified - Begin
            ElseIf l_strName = "JS_BuildInputFile_UM.vbs" Then
              If Not g_dllCmnFunctions.CheckInputFile(g_dicBatchInfo.Item("Database") + ".ecumi.inputfiles",g_dicBatchInfo.Item("WorkfilePath"),g_connDatabase,g_dicBatchInfo) Then
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & " - Error : " & _
                         "eConn UM - No Input Files are available for processing: "
                  g_dicBatchInfo.Item("MESSAGE") = "Log Message : No Input Files are available for processing"
                    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                  g_dicBatchInfo.Item("RETURN CODE") = 0
                  'Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
                  Call f_UnloadGlobalObjects(0)
              End If
              ' 07/08/2010 - Cognizant Offshore Modified - End
            End If
        ' Execution of stored procedure
        ElseIf p_strType = "SqlInfo" Then

            ' Execute the procedure by calling the DLL function ExecuteSP

            If Not g_dllCmnFunctions.ExecuteSP(g_connDatabase,g_dicBatchInfo,l_strName, l_arrParamList,0,l_rstResult) Then

                f_BuildExecuteAction = False
                g_dicBatchInfo.Item("RETURN CODE") = 15
                g_dicBatchInfo.Item("MESSAGE") = "Log Message : Failed to execute the Sql " & l_strName
                wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
            'Syntel - 07/23 : Commented the below lines for 5.2 remediation chagnes;
            'End If
            '
            '04/28/2006 - Cognizant Offshore - Added
            'To check whether the return code is zero or not
            'If cint(l_arrParamList(0, 1)) <> 0  then
            '        g_dicBatchInfo.Item("RETURN CODE") = cint(l_arrParamList(0, 1))
            '        If g_dicBatchInfo.Item("RETURN CODE") = 20 Then
            '            g_dicBatchInfo.Item("MESSAGE") = "Log Message : Bypassed " & g_dicBatchInfo.Item("BatchName") & " Batch"
            '            g_dicBatchInfo.Item("RETURN CODE") = 0
            '           f_BuildExecuteAction = True
            '        Else
            '            f_BuildExecuteAction = False
            '        End If
            '       l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & g_dicBatchInfo.Item("MESSAGE")
            '       wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
                    l_wshShell.LogEvent 0, CStr(Date) & " - " & CStr(Time()) & g_dicBatchInfo.Item("MESSAGE")
                    Call g_dllCmnFunctions.WriteTrace(g_connDatabase,g_dicBatchInfo)
                    Call f_UnloadGlobalObjects(g_dicBatchInfo.Item("RETURN CODE"))
            End If
        End If

        ' Assigning the Message item of the dictionary object
        If l_intReturnCode = 0 Then
            g_dicBatchInfo.Item("MESSAGE") = "Log Message : Successfully executed the " & p_strType
        End If
    Else
        ' Writing the Start of the Build Execute Action Function in Database
        ' Assigning the Message item of the dictionary object
        g_dicBatchInfo.Item("MESSAGE") = "Log Message : Build Input Parameter List for the " & p_strType & "-" & l_strName & " failed"
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
    End If

    ' Destroy the local objects
    Set l_rstResult = Nothing
    Set l_wshShell = Nothing
    Set l_nodeLastChild = Nothing
    Set l_nodeParam = Nothing
    Set l_arrParamList = Nothing

    ' Check for runtime errors
    If Err.Number <> 0 Then
        f_BuildExecuteAction = False
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_BuildExecuteAction Failed; Err Msg : " & Err.Description
    End If

End Function

'05/06/2014 OGS Added Begins
'###########################################################################
'Function Name          : f_Calculate_engique
'Purpose                : 1. To adjust the config as per available claims
'Input                  : gdicoverrideinfo
'Output                 : True/False
'###########################################################################
Private Function f_Calculate_engique()

On Error Resume Next

Dim l_fso
Dim l_regexObj  '//RegExp Object
Dim l_matchObj  '//MatchCollection Object
Dim l_textObj   '//TextStream Object
Dim l_sbufferObj
Dim l_claimperqueue
Dim l_wshShell
dim l_strsql_fce
dim l_database_fce
dim l_batchid_fce
dim l_noofque_fce
dim l_claimcount_fce
dim rstRecordset

    ' Set the function to True by default
    f_Calculate_engique = True


    ' Creating a Shell script Object, to write the error
    Set l_wshShell = wscript.CreateObject("WScript.Shell")
    If Err.Number <> 0 Then
        g_dicBatchInfo.Item("RETURN CODE") = 99
        f_Calculate_engique = False
        Exit Function
    End If

    ' Writing the Start of the Build Execute Action Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Calculating the number of claims Per queue started"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

    l_database_fce = g_dicBatchInfo.Item ("SchemaName")
    l_batchid_fce  = g_dicOverrideInfo.Item ("BatchId")
    l_noofque_fce  = g_dicOverrideInfo.Item ("NumberOfQuesClmu")

    l_strsql_fce   = "SELECT COUNT(1) CLAIM_COUNT FROM " & l_database_fce & ".CMC_CLCL_CLAIM WHERE CLCL_CUR_STS = '16' AND CLCL_BATCH_ID = '" & l_batchid_fce & "';"

    'g_connDatabase.CommandTimeout = 30000
    'Set rstRecordset = g_connDatabase.Execute(l_strsql_fce)
    'Set rstRecordset = g_dllCmnFunctions.

    l_claimcount_fce = rstRecordset("CLAIM_COUNT").Value

    if CInt(l_claimcount_fce) > CInt(l_noofque_fce) then
	l_claimperqueue = CInt(l_claimcount_fce/l_noofque_fce) + 5
	g_dicOverrideInfo.Add "MaxClaimsPerQue",l_claimperqueue
    end if


    ' Writing the End of the Function in Database
    ' Assigning the Message item of the dictionary object
    g_dicBatchInfo.Item("MESSAGE") = "Log Message : Calculating the number of claims Per queue Completed"
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")

' Destroy the local objects
    Set l_fso = Nothing
    Set l_regexObj = Nothing
    Set l_matchObj = Nothing
    Set l_textObj = Nothing
    Set l_sbufferObj = Nothing
    Set l_wshShell= Nothing
    Set l_strsql_fce= Nothing
    Set l_database_fce= Nothing
    Set l_batchid_fce= Nothing
    Set l_noofque_fce= Nothing
    Set l_claimcount_fce= Nothing
    Set l_wshShell = Nothing

  ' Check for runtime errors
    If Err.Number <> 0 Then
        f_BuildExecuteAction = False
        g_dicBatchInfo.Item("RETURN CODE") = 99
        wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - Error : Function f_Calculate_engique Failed; Err Msg : " & Err.Description
    End If
End Function
'05/06/2014 OGS Added ends


'###########################################################################
'Function Name          : f_UnloadGlobalObjects
'Purpose                : 1.To unload all the global objects before exiting
'                         the script.
'Input                  : l_intReturnCode
'Output                 : True/False
'###########################################################################
Private Function f_UnloadGlobalObjects(ByVal p_intReturnCode)
    On Error Resume Next
  If g_dicBatchInfo.Exists("LogFileName") and p_intReturnCode = 0 Then
       Call g_dllCmnFunctions.SendMail(g_connDatabase,g_dicBatchInfo,g_dicBatchInfo.Item("LogFileName"),g_dicBatchInfo.Item("Mail_to_ok"),g_dicBatchInfo.Item("Subject"),g_dicOverrideInfo.Item("RunBookName"))
    g_dicBatchInfo.Item("MESSAGE") = "Log Message :0 " & g_dicBatchInfo.Item("LogFileName") & g_dicOverrideInfo.Item("RunBookName") & p_intReturnCode
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
    ElseIf g_dicBatchInfo.Exists("LogFileName") and p_intReturnCode > 0  Then
    g_dicBatchInfo.Item("MESSAGE") = "Log Message :1 " & g_dicBatchInfo.Item("LogFileName") & g_dicOverrideInfo.Item("RunBookName") & p_intReturnCode
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        Call g_dllCmnFunctions.SendMail(g_connDatabase,g_dicBatchInfo,g_dicBatchInfo.Item("LogFileName"),g_dicBatchInfo.Item("Mail_to_err"),g_dicBatchInfo.Item("Subject"),g_dicOverrideInfo.Item("RunBookName"))
    ElseIf p_intReturnCode > 0 Then
	    g_dicBatchInfo.Item("MESSAGE") = "Log Message : 2 "  & g_dicOverrideInfo.Item("RunBookName") & p_intReturnCode
    wscript.Echo CStr(Date) & " - " & CStr(Time()) & " - " &  g_dicBatchInfo.Item("MESSAGE")
        Call g_dllCmnFunctions.SendMail(g_connDatabase,g_dicBatchInfo,"",g_dicBatchInfo.Item("Mail_to_err"),g_dicBatchInfo.Item("Subject"),g_dicOverrideInfo.Item("RunBookName"))
       End If


    If IsObject(g_dllCmnFunctions) Then
        Set g_dllCmnFunctions = Nothing
    End If

    If IsObject(g_loadXml) Then
        Set g_loadXml = Nothing
    End If

    If IsObject(g_arrPostBypassInfo) Then
        Set g_arrPostBypassInfo = Nothing
    End If

    If IsObject(g_arrPreBypassInfo) Then
        Set g_arrPreBypassInfo = Nothing
    End If

    If IsObject(g_dicOverrideInfo) Then
        Set g_dicOverrideInfo = Nothing
    End If

    If IsObject(g_dicBatchInfo) Then
        Set g_dicBatchInfo = Nothing
    End If

    wscript.Quit (p_intReturnCode)

End Function

