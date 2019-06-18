'#####################################################################################################################
Public Function fnRAD_Login(paramStr)
	CURRENT_FUNCTION = "fnRAD_Login"
	'paramStr = referenceId, [Open new session? - TRUE/FALSE - Default: False]
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", "fnRAD_Login", paramStr)	
	sFlag = False
	indvParam = Split(paramStr, ",")
	referenceId = indvParam(0)
	If UBound(indvParam) > 0 Then		
		newSession = Trim(indvParam(1))
		newSession = TRUE
	Else
		newSession = FALSE
	End If
	
	If Not fnEXCL_setIOReference("AccessDetails.xls", "AccessDetails", referenceId) Then
		Exit Function
	End If
	
	userName = fnEXCL_getValue(INPUT_QRY_REC, "userName")
	userPassword = fnEXCL_getValue(INPUT_QRY_REC, "userPassword")
	accessURL = fnEXCL_getValue(INPUT_QRY_REC, "accessURL")
	Call fnCOM_updateLogFile ("S", "fnRAD_Login", accessURL & "," & userName & "," & userPassword)	
	If newSession Then 
		Call fnCOM_terminateProcess("RADAR")
			Call fnCOM_terminateProcess("radar32")
		SystemUtil.Run accessURL',"","C:\Apps\radar\pre-prod","" -- commented 
		Call fnCOM_SetTestRemarks("Accessed From: " & accessURL)
	ElseIf Not fnCOM_checkProcess("radar32") Then 'fnCOM_checkProcess("RADAR")
		SystemUtil.Run accessURL',"","C:\Apps\radar\pre-prod",""
		Call fnCOM_SetTestRemarks("Accessed From: " & accessURL)
	ElseIf Not PbWindow("PWin_Login").Exist(0) Then
		Call fnRAD_CloseWindow("ALL")
		Call fnCOM_SetTestRemarks("Using open application")
		Call fnCOM_SetTestResult("Pass")
		fnRAD_Login = True
		Exit Function
	End If
	
	lCount = 0 
	Do
		wait(1)
		If fnCOM_checkProcess("radar32") Then ' for acceptance it is "radar32" instead of RADAR which is for Pre-Prod
			If PbWindow("PWin_Login").Exist(0) Then
				Exit Do
			End If
		End If
		lCount = lCount + 1
	Loop While lCount <  100

	PbWindow("PWin_Login").PbEdit("PEdit_UserName").Set userName
	PbWindow("PWin_Login").PbEdit("PEdit_Password").Set userPassword
	Call fnCOM_SetTestRemarks("Logged into application as: " & userName)
	PbWindow("PWin_Login").PbButton("PBtn_Connect").Click
	'wait 2
	lCount = 0 
	Do
		wait(1)
'		' added to avoid manual click on dialog for " message of email id missing"
'		If PbWindow("PWin_Login").Dialog("Dialog_SystemSettings").Exist(0) Then
'			 PbWindow("PWin_Login").Dialog("Dialog_SystemSettings").WinButton("Winbtn_OK").Click
'		End If

'		If Dialog("Dialog_SystemLoginWARNING").Exist(0) Then ' added to click on th ok button for LAN message
'				Dialog("Dialog_SystemLoginWARNING").WinButton("WinBtn_OK").Click		
'		ElseIf 
			IF PbWindow("PWin_MainWindow").Exist(0) Then

			Call fnCOM_SetTestResult("Pass")
			sFlag = True
			Exit Do
		
		ElseIf Dialog("Dialog_SystemLoginWARNING").Exist(0) Then
			If InStr(1, Dialog("Dialog_SystemLoginWARNING").GetROProperty("text"), "System") Then
				Call fnCOM_SetTestResult("Pass")
				sFlag = True
				Dialog("Dialog_SystemLoginWARNING").WinButton("WinBtn_OK").Click ' moved from line 85 to include it inside the if condition
				wait 7
			Else
				Call fnCOM_SetTestResult("Fail")
				sFlag = False
				errStr = Dialog("PDia_Alert").Static("WinTxt_Alert").GetROProperty("text")
				Call fnCOM_SetTestRemarks("[ERR] Issue: " & errStr)
			End If

			Dialog("PDia_Alert").WinButton("WinBtn_OK").Click
			wait 7
			Exit Do
		End If
		lCount = lCount + 1
	Loop While lCount < 300
	
	If not sFlag Then
		Call fnCOM_SetTestRemarks("[ERR] Issues in logging into the app")
		Call fnCOM_SetTestResult("Fail")
	End If
	
	If Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(10) Then
		Dialog("PDia_Alert").WinButton("WinBtn_OK").Click
	End If
	
	fnRAD_CloseWindow("MIM Alert Log")
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_SetPopDataValue(paramStr)
	'paramStr = objectName, [objectValue = xyz or <blank> or <rand> or ""]
	CURRENT_FUNCTION = "fnRAD_SetPopDataValue"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, paramStr)	
	sFlag = False
	indvParam = Split(paramStr, ",")
	If UBound(indvParam) < 2 Then
		Call fnCOM_updateLogFile ("E", CURRENT_FUNCTION, "Invalid Parameters:" & paramStr)
	End If
	SubwindowName = Trim(indvParam(0))
	dwName = Trim(indvParam(1))
	Rowvalue=Trim(indvParam(2))
	objectName = Trim(indvParam(3))
	objectValue = Trim(indvParam(4))	
	
	If Left(objectValue, 2) = "c_" Then		
		colName = Replace(objectValue, "c_", "")		
		outputUpdt = False
		objectValue = INPUT_QRY_REC.Fields.Item(colName)
		If InStr(1, objectValue, "<rand>", 1) Then outputUpdt = True
	End If
	Call fnCOM_updateLogFile("I", CURRENT_FUNCTION, objectName & "=" & objectValue)	
	
	If PbWindow("PWin_MainWindow").Exist(10) Then
		If objectValue <> "" Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow(SubwindowName).PbDataWindow(dwName)						
			If rootObj.Exist(0) Then	
				
				objectValue = Replace(objectValue, "<rand>", Left(Hex(Timer) & Hex(Date), 8))
				If IsDate(objectValue) Then
					objectValue = Day(objectValue) & "/" & Month(objectValue) & "/" & Right(Year(objectValue), 2)
				End If				
				rootObj.SetCellData Rowvalue, objectName, objectValue	
					
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks(objectName & " set to " & objectValue)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objectName)
			End IF
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
		End If	
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_SetObjValue(paramStr)
	'paramStr = objectName, [objectValue = xyz or <blank> or <rand> or ""]
	CURRENT_FUNCTION = "fnRAD_SetObjValue"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, paramStr)
	
	sFlag = False
	indvParam = Split(paramStr, ",")
	If UBound(indvParam) < 2 Then
		Call fnCOM_updateLogFile ("E", CURRENT_FUNCTION, "Invalid Parameters:" & paramStr)
	End If
	windowName = Trim(indvParam(0))
	objectName = Trim(indvParam(1))
	objectValue = Trim(indvParam(2))
	
	If Left(objectValue, 2) = "c_" Then
		colName = Replace(objectValue, "c_", "")
		outputUpdt = False
		objectValue = INPUT_QRY_REC.Fields.Item(colName)
		'If objectValue= "Navion Oslo" then		 			
		'End If	
		If InStr(1, objectValue, "<rand>", 1) Then outputUpdt = True
	End If
	
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If objectValue <> "" Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow(windowName)			
			If rootObj.Exist(0) Then
				objectValue = Replace(objectValue, "<rand>", Left(Hex(Timer) & Hex(Date), 8))
				If IsDate(objectValue) Then
					objectValue = Day(objectValue) & "/" & Month(objectValue) & "/" & Right(Year(objectValue), 2)
				End If		   					
				rootObj.PbEdit(objectName).Set objectValue									
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks(objectName & " set to " & objectValue)
'		ElseIf RADAR_DEAL_TYPE="VCO" Then
'				Set dwObj=PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_Deal")
'					 FOR li_Loop = 1 TO rootObj.object.datawindow.column.count
'					    ls_ColName =colName
'					    IF  rootObj.Describe( ls_ColName + ".Visible") > 0 THEN
'					    dwObj.SelectCell "#1",li_Loop,objectValue
'					    END IF					
'					 next
'					Call fnCOM_SetTestResult("Pass")
'					Call fnCOM_SetTestRemarks(objectName & " set to " & objectValue)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objectName)
			End IF
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
		End If	
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
	End If	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_ClickObject(objName)
	CURRENT_FUNCTION = "fnRAD_ClickObject"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, objName)	
	fnRAD_ClickObject = False
	
	
	'***********************************************************************************************
	'Added on 08-04-2019 to include Ok Button when a port leg is added to the voyage and the "PortcostLookup" dialog appears
	If PbWindow("PWin_MainWindow").Dialog("PDia_Portcostlookup").Exist(2) Then
				PbWindow("PWin_MainWindow").Dialog("PDia_Portcostlookup").WinButton("WinBtn_OK").Click	
				Wait 1
	End If
	'***********************************************************************************************
	
	
	If PbWindow("PWin_MainWindow").Exist(10) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(10) Then			
			If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton(objName).Exist(10) Then
				PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton(objName).Click
				'wait 2
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks("Clicked on: " & objName)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objName)
			End If
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(2) Then			
			If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbButton(objName).Exist(2) Then
				PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbButton(objName).Click
				'wait 2
				If objName = "PBtn_MVoyage" Then
					If fnRAD_VerifyTitle("Vessel Search") Then
						Call fnRAD_CloseWindow("Vessel Search") 
					End If
				End If				
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks("Clicked on: " & objName)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objName)
			End If
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objName)
		End If
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objName)
	End If
	
	lCount = 0
	Do
		wait(2)		
		If Window("Win_RADAR").WinMenu("Win_Menu").Exist(20) Then
			Exit Do
		End If
		lCount = lCount + 1
	Loop While lCount < 5
	
	
	If instr(objName, "PBtn_Claims")<>0 or instr(objName, "PBtn_MDealEntry")<>0 Then
		'Call fnRAD_WaitforReady()
		Call fnRAD_WaitUntilWinLoads()
		'wait 180
		wait 20
	End If
	
	Call fnRAD_WaitUntilWinLoads()
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_NavigateTo(strMenu)
	CURRENT_FUNCTION = "fnRAD_NavigateTo"
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, strMenu)	
	If Window("Win_RADAR").Exist(10) Then
		On Error Resume Next
		Window("Win_RADAR").WinMenu("Win_Menu").Select strMenu
		'wait(2)
		If Err.Number <> 0 Then
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Unable to navigate to: " & strMenu)
		Else
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Navigated To: " & strMenu)
		End If
	Else 
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Unable to navigate to: " & strMenu)
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_CloseWindow(paramStr)
	CURRENT_FUNCTION = "fnRAD_CloseWindow"
	On Error Resume Next
	'paramStr = windowName/All
	'wait 2
	Set oDesc = Description.Create()
	oDesc("nativeclass").Value = "FNWND3125"
	Set allWindows = PbWindow("PWin_MainWindow").ChildObjects(oDesc)
	For wCount  = 0 To allWindows.Count - 1
		winTitle = allWindows(wCount).GetROProperty("text")
		If InStr(1, winTitle, "Accelerator", 1) < 1 Then
			If paramStr = "" Or paramStr = "ALL" Then 	
				'wait 2
				Call fnRAD_HandleDialog()
				allWindows(wCount).Close
				Call fnRAD_HandleDialog()
				wait 1
				Call fnCOM_updateLogFile("I", "Closed Window: ", winTitle)
				'Call fnCOM_SetTestResult("Pass")
			ElseIf InStr(1, winTitle, paramStr, 1) > 0 Then
				'wait 2
				Call fnRAD_HandleDialog()
				allWindows(wCount).Close
				Call fnRAD_HandleDialog()
				wait 1
				Call fnCOM_updateLogFile("I", "Closed Window: ", winTitle)
				'Call fnCOM_SetTestResult("Pass")
			End If
		End If	
	Next
	Set oDesc = Description.Create()
	oDesc("nativeclass").Value = "FNWNS3125"
	Set allWindows = PbWindow("PWin_MainWindow").ChildObjects(oDesc)
	For wCount  = 0 To allWindows.Count - 1
		winTitle = allWindows(wCount).GetROProperty("text")
		If InStr(1, winTitle, "Accelerator", 1) < 1 Then
			If paramStr = "" Or paramStr = "ALL" Then 
				'wait 2
				allWindows(wCount).Close
				fnRAD_HandleDialog()
				wait 1
				Call fnCOM_updateLogFile("I", "Closed Window: ", winTitle)
			ElseIf InStr(1, winTitle, paramStr, 1) > 0 Then
				'wait 2
				allWindows(wCount).Close
				fnRAD_HandleDialog()
				wait 1
				Call fnCOM_updateLogFile("I", "Closed Window: ", winTitle)
			End If
		End If	
	Next
	Call fnCOM_SetTestResult("Pass")
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

Function fnRAD_HandleDialog()
	If PbWindow("PWin_MainWindow").PbWindow("text:=Trip Vetting Explorer").Exist(2) Then
		PbWindow("PWin_MainWindow").PbWindow("text:=Trip Vetting Explorer").Close
		Dialog("PDia_Alert").WinButton("WinBtn_Yes").Click
	End If
	
	If Dialog("PDia_Alert").WinButton("WinBtn_No").Exist(2) Then
		Dialog("PDia_Alert").WinButton("WinBtn_No").Click
	End If
	
	If Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(2) Then
		Dialog("PDia_Alert").WinButton("WinBtn_OK").Click
	End If
End Function

'#####################################################################################################################
Public Function fnRAD_WaitForWindow(waitTitle)
	CURRENT_FUNCTION = "fnRAD_WaitForWindow"
	On Error Resume Next
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, waitTitle)
	
	sFlag = False
	lCount = 0
	
	Do
		wait(2)
		actTitle = PbWindow("PWin_MainWindow").GetROProperty("text")
		If actTitle <> "" And (InStr(1, actTitle, waitTitle, 1) Or InStr(1, waitTitle, actTitle, 1)) Then
			sFlag = True
		Else
			Set oDesc = Description.Create()
			oDesc("nativeclass").Value = "FNWND3125"
			Set allWindows = PbWindow("PWin_MainWindow").ChildObjects(oDesc)
			For wCount  = 0 To allWindows.Count - 1
				winTitle = allWindows(wCount).GetROProperty("text")			
				If winTitle <> "" And (InStr(1, winTitle, waitTitle, 1) Or InStr(1, waitTitle, winTitle, 1)) Then
					sFlag = True
					Exit For
				End If
			Next
		End If	
		If Not sFlag Then
			Set oDesc = Description.Create()
			oDesc("nativeclass").Value = "FNWNS3125"
			Set allWindows = PbWindow("PWin_MainWindow").ChildObjects(oDesc)
			For wCount  = 0 To allWindows.Count - 1
				winTitle = allWindows(wCount).GetROProperty("text")
				If winTitle <> "" And (InStr(1, winTitle, waitTitle, 1) Or InStr(1, waitTitle, winTitle, 1)) Then
					sFlag = True
					Exit For
				End If
			Next
		End If
		lCount = lCount + 1
		If lCount > 6 Then
			Exit Do
		End If
	Loop While sFlag = False
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_VerifyTitle(expTitle)
	CURRENT_FUNCTION = "fnRAD_VerifyTitle"
	On Error Resume Next
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, expTitle)
	'wait 2
	Call fnRAD_WaitUntilWinLoads()
	sFlag = False
	
	actTitle = PbWindow("PWin_MainWindow").GetROProperty("text")
	
'	If actTitle <> "" And (InStr(1, actTitle, expTitle, 1) Or InStr(1, expTitle, actTitle, 1)) Then
'		sFlag = True
'		resultStr = "Pass"
'		remarksStr = "Title: " & actTitle
'	Else
'		Set oDesc = Description.Create()
'		oDesc("nativeclass").Value = "FNWND3125"
'		Set allWindows = PbWindow("PWin_MainWindow").ChildObjects(oDesc)
'		For wCount  = 0 To allWindows.Count - 1
'			winTitle = allWindows(wCount).GetROProperty("text")								
'			If winTitle <> "" And (InStr(1, winTitle, expTitle, 1) Or InStr(1, expTitle, winTitle, 1)) Then
'				sFlag = True
'				resultStr = "Pass"
'				remarksStr = "Title: " & winTitle
'				Exit For
'			End If			
'		Next
'	End If	
'	If Not sFlag Then
'		Set oDesc = Description.Create()
'		oDesc("nativeclass").Value = "FNWNS3125"
'		Set allWindows = PbWindow("PWin_MainWindow").ChildObjects(oDesc)
'		For wCount  = 0 To allWindows.Count - 1
'			winTitle = allWindows(wCount).GetROProperty("text")			
'			If winTitle <> "" And (InStr(1, winTitle, expTitle, 1) Or InStr(1, expTitle, winTitle, 1)) Then
'				sFlag = True
'				resultStr = "Pass"
'				remarksStr = "Title: " & winTitle
'				Exit For
'			End If
'		Next	
'	End If
'	
'	If Not sFlag Then
'		resultStr = "Fail"
'		remarksStr = "[ERR] No window found with title: " & expTitle
'	End If

	If actTitle <> "" And (InStr(1, actTitle, expTitle, 1) Or InStr(1, expTitle, actTitle, 1)) Then
		sFlag = True
		resultStr = "Pass"
		remarksStr = "Title: " & actTitle
	Else
		For i=1 to 50
			Set oDesc = Description.Create()
			oDesc("nativeclass").Value = "FNWN.*"
			Set allWindows = PbWindow("PWin_MainWindow").ChildObjects(oDesc)
			For wCount  = 0 To allWindows.Count - 1
				winTitle = allWindows(wCount).GetROProperty("text")								
				If winTitle <> "" And (InStr(1, winTitle, expTitle, 1) Or InStr(1, expTitle, winTitle, 1)) Then
					sFlag = True
					resultStr = "Pass"
					remarksStr = "Window Displayed: " & winTitle
					Exit For
				End If			
			Next
			If sFlag Then
				Exit For
			End If
		Next
		End If	
		
		If Not sFlag Then
			resultStr = "Fail"
			remarksStr = "[ERR] No window found with title: " & expTitle
		End If
		
		If instr(expTitle, "Receive Mail")>0 Then
			If Dialog("PDia_Alert").WinButton("WinBtn_Cancel").Exist(5) Then
				Dialog("PDia_Alert").WinButton("WinBtn_Cancel").Click
				wait 5
			End If
			If Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(5) Then
 				fnRAD_ClickAlert("WinBtn_OK")
			End If
		End If

	fnRAD_VerifyTitle = sFlag
	Call fnCOM_SetTestResult(resultStr)
	Call fnCOM_SetTestRemarks(remarksStr)
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_DblClickRow(paramStr)
	CURRENT_FUNCTION = "fnRAD_DblClickRow"	
	On Error Resume Next
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, rowStr)	
	
	indvParam = Split(paramStr, ",")
	If UBound(indvParam) < 1 Then
		Call fnCOM_updateLogFile ("E", CURRENT_FUNCTION, "Invalid Parameters:" & paramStr)
		Exit Function
	End If
	rowStr = Trim(indvParam(0))
	SubwindowName = Trim(indvParam(1))
	
	sFlag = False
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbDataWindow(SubwindowName).Exist(0) Then
				Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbDataWindow(SubwindowName)			
				sFlag = True
			End If
		End If
		If sFlag = False Then
			If  PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
				If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow(SubwindowName).Exist(0) Then
					Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow(SubwindowName)			
					sFlag = True
				End If
			End If
		End If
		If sFlag = False Then
			Call fnCOM_SetTestResult("Fail")
            Call fnCOM_SetTestRemarks("[ERR] Window unavailable")
			Exit Function
		End If
	End If	
	If sFlag Then
		sFlag = False
		totalRows = rootObj.RowCount		
		totalCols = rootObj.ColumnCount		
		For cCount = 1 To totalCols
			For rCount = 1 To totalRows
				rValue = rootObj.GetCellData(rCount, cCount)						
				Call fnCOM_updateLogFile("I", "Data lookup", rValue & " found at " & rCount & "," & cCount)		
				If InStr(1, rValue, rowStr, 1) Then		
					rootObj.SelectCell rCount, cCount
					rootObj.ActivateCell rCount, cCount
					wait 1
					Call fnCOM_updateLogFile("I", CURRENT_FUNCTION, rowStr & " found at " & rCount & "," & cCount)		
					sFlag = True	
					Exit For
				End If
			Next
			If sFlag Then 
				resultStr = "Pass"
				remarksStr = "Selected: '" & rowStr & "'"
				Exit For
			End If
		Next
		If Not sFlag Then
			resultStr = "Fail"
			remarksStr = "[ERR] '" & rowStr & "' not found"
		End If
	End If
	
	fnRAD_DblClickRow = sFlag
	Call fnCOM_SetTestResult(resultStr)
	Call fnCOM_SetTestRemarks(remarksStr)
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_SelectRow(paramStr)
	CURRENT_FUNCTION = "fnRAD_SelectRow"
	On Error Resume Next
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, paramStr)
	
	tStr = Split(paramStr, ",")
	rowStr = Trim(tStr(0))
	If UBound(tStr) > 0 Then
		tableName = Trim(tStr(1))
	Else
		tableName = "PDw_DataWindow"
	End If
	
	sFlag = False
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbWindow("PWin_SubWin2_1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbWindow("PWin_SubWin2_1")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")
			sFlag = True
		Else
			resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
	End If		
	If sFlag Then
		sFlag = False
		totalRows = rootObj.PbDataWindow(tableName).RowCount
		totalCols = rootObj.PbDataWindow(tableName).ColumnCount
		For cCount = 1 To totalCols
			For rCount = 1 To totalRows
				rValue = rootObj.PbDataWindow(tableName).GetCellData(rCount, cCount)					
				If InStr(1, UCase(rValue),UCase(rowStr), 1) Then										
					rootObj.PbDataWindow(tableName).ActivateCell rCount, cCount
					Call fnCOM_updateLogFile("I", CURRENT_FUNCTION, rowStr & " found at " & rCount & "," & cCount)
					sFlag = True
					If rootObj.PbButton("PBtn_Choose").Exist(0) Then
						rootObj.PbButton("PBtn_Choose").Click
						wait 1
					End If
					Exit For
				End If
			Next			
			If sFlag Then 
				resultStr = "Pass"
				remarksStr = "Selected: '" & rowStr & "'"
				Call fnRAD_CloseAlerts()
				Exit For
			End If
		Next
		If Not sFlag Then
			resultStr = "Fail"
			remarksStr = "[ERR] '" & rowStr & "' not found"
		End If
	End If
	
	fnRAD_SelectRow = sFlag
	Call fnCOM_SetTestResult(resultStr)
	Call fnCOM_SetTestRemarks(remarksStr)
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_SelectVoyage(paramStr)
	CURRENT_FUNCTION = "fnRAD_SelectVoyage"
	On Error Resume Next
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, paramStr)
	
	tStr = Split(paramStr, ",")
	rowStr = Trim(tStr(0))
	If UBound(tStr) > 0 Then
		tableName = Trim(tStr(1))
	Else
		tableName = "PDw_DataWindow"
	End If
	
	sFlag = False
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")
			sFlag = True
		Else
			resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
	End If	
	If sFlag Then
		sFlag = False
		totalRows = rootObj.PbDataWindow(tableName).RowCount
		totalCols = rootObj.PbDataWindow(tableName).ColumnCount
		'For cCount = 1 To totalCols
			For rCount = 1 To totalRows
				rValue = rootObj.PbDataWindow(tableName).GetCellData(rCount, 5)				
				If rValue = rowStr Then					
					'rootObj.PbDataWindow(tableName).SelectCell rCount, cCount
					rootObj.PbDataWindow(tableName).ActivateCell rCount, "#5"
					'Call fnCOM_updateLogFile("I", CURRENT_FUNCTION, rowStr & " found at " & rCount & "," & cCount)
					sFlag = True					
					Exit For
				End If
			Next
			If sFlag Then 
				resultStr = "Pass"
				remarksStr = "Selected: '" & rowStr & "'"
				Call fnRAD_CloseAlerts()
				'Exit For
			End If
		'Next
		If Not sFlag Then
			resultStr = "Fail"
			remarksStr = "[ERR] '" & rowStr & "' not found"
		End If
	End If
	
	fnRAD_SelectRow = sFlag
	Call fnCOM_SetTestResult(resultStr)
	Call fnCOM_SetTestRemarks(remarksStr)
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function


'#####################################################################################################################
Public Function fnRAD_CloseAlerts()
	CURRENT_FUNCTION = "fnRAD_CloseAlerts"
	On Error Resume Next
	
	aCount = 1
	Do While Dialog("PDia_Alert").Exist(5) 
		alertTxt = Dialog("PDia_Alert").Static("WinTxt_Alert").GetROProperty("text")
		If Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(0) Then
			Dialog("PDia_Alert").WinButton("WinBtn_OK").Click
		ElseIf Dialog("PDia_Alert").WinButton("WinBtn_Yes").Exist(0) Then
			Dialog("PDia_Alert").WinButton("WinBtn_Yes").Click
		End If
		wait 1
		Call fnCOM_updateLogFile("I", "Closed alert", alertTxt)
		Call fnCOM_SetTestRemarks("Alert " & aCount & ": " & alertTxt)
		aCount = aCount + 1
		If aCount > 10 Then
			Exit Do
		End If
	Loop
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_VerifyAlert(verifyTxt)
	CURRENT_FUNCTION = "fnRAD_VerifyAlert"
	On Error Resume Next
	
	If Dialog("PDia_Alert").Exist(10) Then
		alertTxt = Dialog("PDia_Alert").Static("WinTxt_Alert").GetROProperty("text")
		If InStr(1, alertTxt, verifyTxt, 1) Then
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Alert: " & alertTxt)
			fnRAD_VerifyAlert = True
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Actual Alert: " & alertTxt)
		End If
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Alert not displayed")
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_ClickAlert(alertBtn)
	CURRENT_FUNCTION = "fnRAD_ClickAlert"
	On Error Resume Next
	'wait 3
	'msgbox "Chk"
	If Dialog("PDia_Alert").Exist(10) Then
		If Dialog("PDia_Alert").WinButton(alertBtn).Exist(10) Then		
			Dialog("PDia_Alert").WinButton(alertBtn).Click
			wait 1
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Clicked On: " & alertBtn)			
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Button not found: " & alertBtn)
		End If
	ElseIf PbWindow("PWin_MainWindow").Dialog("PDia_VieweRDA").Exist(10) Then     'Added on 28022019 to include click of the View eRDA "OK" button
		If PbWindow("PWin_MainWindow").Dialog("PDia_VieweRDA").WinButton(alertBtn).Exist(10) Then		
			PbWindow("PWin_MainWindow").Dialog("PDia_VieweRDA").WinButton(alertBtn).Click
			wait 1
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Clicked On: " & alertBtn)			
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Button not found: " & alertBtn)
		End If
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Alert not displayed")
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_SelToolBar(clickOpt)
	CURRENT_FUNCTION = "fnRAD_SelToolBar"
	On Error Resume Next
	
	Call fnRAD_WaitUntilWinLoads()
	If PbWindow("PWin_MainWindow").PbObject("PObj_ToolBar").Exist(0) Then
		stoolbartxt=PbWindow("PWin_MainWindow").PbObject("PObj_ToolBar").GetVisibleText	
		If instr(stoolbartxt, "Voyage")=0 Then
			PbWindow("PWin_MainWindow").PbObject("PObj_ToolBar").Click 10, 10, micRightBtn
			'Window("Win_RADAR").WinMenu("Win_Voyage").Select "Show Text"
			Window("Win_RADAR").WinMenu("Win_RightClickMenu").Select "Show Text"
			wait 3
		End If		

		PbWindow("PWin_MainWindow").PbObject("PObj_ToolBar").GetTextLocation clickOpt, xLeft, xTop, xRight, xBottom
		If xRight > 4 Then			
			PbWindow("PWin_MainWindow").PbObject("PObj_ToolBar").Click xRight ,xTop			
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Clicked on: " & clickOpt)
			wait 1
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Option not found: " & clickOpt)
		End If
'		Select Case clickOpt
'			Case "Find"
			
'				If PbWindow("PWin_MainWindow").InsightObject("IO_Find").Exist Then
'					PbWindow("PWin_MainWindow").InsightObject("IO_Find").Click
'				Else
'					resultstr = "Fail"
'				End If
'				
'			Case "Insert"
'				If	PbWindow("PWin_MainWindow").InsightObject("IO_Insert").Exist Then
'					PbWindow("PWin_MainWindow").InsightObject("IO_Insert").Click
'				Else
'					resultstr = "Fail"
'				End If
'		End Select
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Toolbar not visible")
	End If
	'wait 2
	
'	If resultstr = "Fail" Then
'		Call fnCOM_SetTestRemarks("[ERR] "&clickOpt&" not found")
'	End If

	If instr(clickOpt, "Deal")<>0 Then
		wait 30
		Call fnRAD_WaituntilWinLoads()
	End If

	If instr("Save", clickOpt)<>0 Then
		If Dialog("PDia_Alert").WinButton("WinBtn_Yes").Exist(5) Then Call fnRAD_ClickAlert("WinBtn_Yes")
	End If

	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_ClickPbObj(clickOpt)
	CURRENT_FUNCTION = "fnRAD_ClickPbObj"
	
	Select Case clickOpt
	
	Case "Insert"
		xRight="81"
		xTop= "15"
	Case "Find"
		xRight="225"
		xTop= "15"	
	End select 
	
	On Error Resume Next	
	
	If PbWindow("PWin_MainWindow").PbObject("PObj_ToolBar").Exist(0) Then				
		PbWindow("PWin_MainWindow").PbObject("PObj_ToolBar").Click xRight , xTop			
		Call fnCOM_SetTestResult("Pass")
		Call fnCOM_SetTestRemarks("Clicked on: " & clickOpt)	
		wait 1	
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Toolbar not visible")
	End If
	'wait 2
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_VerifyAndClose(expTitle)
	CURRENT_FUNCTION = "fnRAD_VerifyAndClose"
	On Error Resume Next
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, expTitle)
	
	Call fnRAD_VerifyTitle(expTitle)
	Call fnRAD_CloseWindow(expTitle)
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'############################################################################################################################
Public Function fnRAD_SetDataValue(paramStr)
	'paramStr = objectName, [objectValue = xyz or <blank> or <rand> or ""]
	CURRENT_FUNCTION = "fnRAD_SetDataValue"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, paramStr)
	
	sFlag = False
	indvParam = Split(paramStr, ",")
	If UBound(indvParam) < 2 Then
		Call fnCOM_updateLogFile ("E", CURRENT_FUNCTION, paramStr)
	End If
	windowName = Trim(indvParam(0))
	rownNo=Trim(indvParam(1))
	objectName = Trim(indvParam(2))
	strObjValue = Trim(indvParam(3))
	
	If instr(rownNo,"#")=0 Then
		rownNo = "#"&rownNo
	End If
	
	If Left(strObjValue, 2) = "c_" Then		
		colName = Replace(strObjValue, "c_", "")			
		outputUpdt = False		
		strObjValue = INPUT_QRY_REC.Fields.Item(colName)		
		If colName="stype" Then
			If strObjValue="VCO" Then
				RADAR_DEAL_TYPE=strObjValue
			End If
		End If
		If InStr(1, strObjValue, "<rand>", 1) Then outputUpdt = True
	End If		
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")
			sFlag = True
		
		ElseIf PbWindow("PWin_MainWindow").PbWindow("w_upd_trip").Exist(0) Then
			Set rootObj=PbWindow("PWin_MainWindow").PbWindow("w_upd_trip")
		EsleIf
			resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
	End If	
	
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If strObjValue <> "" Then
			If rootObj.Exist(20) Then
				strObjValue = Replace(strObjValue, "<rand>", Left(Hex(Timer) & Hex(Date), 8))	
				If IsDate(strObjValue) Then
					strObjValue = Day(strObjValue) & "/" & Month(strObjValue) & "/" & Right(Year(strObjValue), 2)					
				End If
				rootObj.PbDataWindow(windowName).SetCellData  rownNo, objectName, strObjValue				
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks(objectName & " set to " & strObjValue)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objectName)
			End IF
		End If	
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'############################################################################################################################
Public Function fnRAD_SelectDataValue(paramStr)
	'paramStr = objectName, [objectValue = xyz or <blank> or <rand> or ""]
	CURRENT_FUNCTION = "fnRAD_SelectDataValue"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, paramStr)
	
	sFlag = False
	indvParam = Split(paramStr, ",")
	
	windowName = Trim(indvParam(0))
	rownNo=Trim(indvParam(1))
	objectName = Trim(indvParam(2))
	'strObjValue = Trim(indvParam(2))		
	
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")			
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")			
			sFlag = True
		Else
			resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
	End If
	
	If PbWindow("PWin_MainWindow").Exist(10) Then
		'Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWindow").PbDataWindow(windowName)
		If objectName <> "" Then
			If rootObj.Exist(10) Then					
				rootObj.PbDataWindow(windowName).SelectCell rownNo, objectName
				rootObj.PbDataWindow(windowName).ActivateCell rownNo, objectName					
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks(objectName & " set to " & objectName)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objectName)
			End IF
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
		End If	
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_ClickMObject(objName)
	CURRENT_FUNCTION = "fnRAD_ClickMObject"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, objName)	
	fnRAD_ClickMObject = False
	
	Call fnRAD_CloseWindow("ALL")
	If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton(objName).Exist(10) Then		
		PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton(objName).Click
		'wait 3	
		Call fnCOM_SetTestResult("Pass")
		Call fnCOM_SetTestRemarks("Clicked on: " & objName)
	Else
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbButton(objName).Exist(10) Then				 	
				PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbButton(objName).Click
				'wait 3
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks("Clicked on: " & objName)
		Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objName)
		End If
	End If
	
	lCount = 0
	Do
		wait(2)		
		If Window("Win_RADAR").WinMenu("Win_Menu").Exist(20) Then
			Exit Do
		End If
		lCount = lCount + 1
	Loop While lCount < 5
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_TypeDataValues(paramStr)
	'paramStr = objectName, [objectValue = xyz or <blank> or <rand> or ""]
	CURRENT_FUNCTION = "fnRAD_TypeDataValues"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, paramStr)
	
	sFlag = False
	indvParam = Split(paramStr, ",")
	If UBound(indvParam) < 2 Then
		Call fnCOM_updateLogFile ("E", CURRENT_FUNCTION, paramStr)
	End If
	windowName = Trim(indvParam(0))
	objectName = Trim(indvParam(1))
	strObjValue = Trim(indvParam(2))
		
	If Left(strObjValue, 2) = "c_" Then		
		colName = Replace(strObjValue, "c_", "")		
		outputUpdt = False
		strObjValue = INPUT_QRY_REC.Fields.Item(colName)
		If InStr(1, strObjValue, "<rand>", 1) Then outputUpdt = True
	End If
	
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")
			sFlag = True
		Else
		resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
		If PbWindow("PWin_MainWindow").PbWindow("wr_define_params_dem").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("wr_define_params_dem")
			rootObj.highlight
			sFlag = True  
		else
		resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
		
		
	End If
	
	
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If strObjValue <> "" Then
			If rootObj.Exist(0) Then
				strObjValue = Replace(strObjValue, "<rand>", Left(Hex(Timer) & Hex(Date), 8))	
				If IsDate(strObjValue) Then
					strObjValue = Day(strObjValue) & "/" & Month(strObjValue) & "/" & Right(Year(strObjValue), 2)					
				End If
				rootObj.PbDataWindow(windowName).SelectCell "#1", objectName
				rootObj.PbDataWindow(windowName).Type strObjValue	
				wait 1	
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks(objectName & " set to " & strObjValue)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objectName)
			End IF
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
		End If	
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function
'#####################################################################################################################
Public Function fnRAD_FormatDate(strparam)
	CURRENT_FUNCTION = "fnRAD_FormatDate"
	On Error Resume Next
	Call fnCOM_updateLogFile ("I", CURRENT_FUNCTION, paramStr)

indvParam=split(strparam,",")
dDate=indvParam(0)
format=indvParam(1)

Select Case format
	Case "DDMMMYYYY"
		sFormat=FormatDateTime(dDate, 1)
		spldate=split(sFormat, " ")
		newFormat=spldate(0)&ucase(left(spldate(1),3))&spldate(2)
		'Sanghamitra
		spldate=split(dDate,"-")
       mon=left(monthname(month(dDate)),3)
      newformat=spldate(0)&mon&"20"&spldate(2)
	  
	Case "DDMMMYY"
		sFormat=FormatDateTime(dDate, 1)
		spldate=split(sFormat, " ")
		newFormat=spldate(0)&ucase(left(spldate(1),3))&right(spldate(2),2)
	Case "MMMYY"
		sFormat=FormatDateTime(dDate, 1)
		spldate=split(sFormat, " ")
		newFormat=right(spldate(2),2)&ucase(left(spldate(1),3))
End Select
	
fnRAD_FormatDate=newFormat

End Function
'#####################################################################################################################
Public Function fnRAD_VerifyDialog(str)
	CURRENT_FUNCTION = "fnRAD_VerifyDialog"
	On Error Resume Next
	Call fnCOM_updateLogFile ("I", CURRENT_FUNCTION, paramStr)
	wait(5)
	fnRAD_VerifyDialog=False
	'sanghamitra
	'	If Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(1) Then
	'	 Dialog("PDia_Alert").WinButton("WinBtn_OK").Click
		If Dialog("PDia_Alert").WinButton("WinBtn_Yes").Exist(1) Then
			Dialog("PDia_Alert").WinButton("WinBtn_Yes").Click
 		End If
	If Dialog("PDia_Alert").Static("WinTxt_Alert").Exist(5) Then
		sText=trim(Dialog("PDia_Alert").Static("WinTxt_Alert").GetROProperty("text"))
		If instr(sText, str)<>0 Then
			resultStr="Pass"
			fnRAD_VerifyDialog=True
			remarksStr="Dialog displayed: "&str
		Else
			resultStr="Fail"
			remarksStr="[ERR] "&sText			
		End If
		
		If Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(1) Then
			Dialog("PDia_Alert").WinButton("WinBtn_OK").Click
		ElseIf Dialog("PDia_Alert").WinButton("WinBtn_Yes").Exist(1) Then
			Dialog("PDia_Alert").WinButton("WinBtn_Yes").Click
 		End If
 		
 		If Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(1) Then
			Dialog("PDia_Alert").WinButton("WinBtn_OK").Click
		ElseIf Dialog("PDia_Alert").WinButton("WinBtn_Yes").Exist(1) Then
			Dialog("PDia_Alert").WinButton("WinBtn_Yes").Click
 		End If
 	Else
 		resultStr="Fail"
		remarksStr="[ERR] No dialog displayed"	
	End If

	Call fnCOM_SetTestResult(resultStr)
	Call fnCOM_SetTestRemarks(remarksStr)
	Call fnCOM_reportErr(CURRENT_FUNCTION)	
End Function

'#####################################################################################################################
Public Function fnRAD_RightClickNavigateTo(strmenu)
	CURRENT_FUNCTION = "fnRAD_RightClickNavigateTo"
	On Error Resume Next
	Call fnCOM_updateLogFile("I", CURRENT_FUNCTION, strMenu)
	wait 2
	strObject= Split(strMenu, ",")
	strobj=strObject(0)
	menuItem = Split(strObject(1), ";")
	xright = strObject(2)
	xbtm = strObject(3)
	err.clear
	
	Select Case strobj
		Case "DEAL"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_Cargo").Click 354, 98, micRightBtn
			wait 2
				If InStr(1, "Initiate eRDA", trim(menuItem(0)), 1) Then
					levelNo1 = 7
				ElseIf InStr(1, "View eRDA", trim(menuItem(0)), 1) Then
					levelNo1 = 8
				ElseIf InStr(1, "Correspondence", trim(menuItem(0)), 1) Then
					levelNo1 = 9
				ElseIf InStr(1, "Performance Indicators", trim(menuItem(0)), 1) Then
					levelNo1 = 11
				End If
				
		Case "Deal"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_Deal").Click 354, 98, micRightBtn
			wait 2
				If InStr(1, "Copy Deal", trim(menuItem(0)), 1) Then
					levelNo1 = 1
				End If

		Case "CARGO"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_Cargo").Click 354, 98, micRightBtn
			wait 2
				If InStr(1, "Performance Indicators", trim(menuItem(0)), 1) Then
					levelNo1 = 16
				ElseIf InStr(1, "Laytime Options", trim(menuItem(0)), 1) Then
					levelNo1 = 5
				ElseIf InStr(1, "Cargo Notes", trim(menuItem(0)), 1) Then
					levelNo1 = 6
				ElseIf InStr(1, "VCI Link", trim(menuItem(0)), 1) Then
					levelNo1 = 10
				ElseIf InStr(1, "Summary", trim(menuItem(0)), 1) Then
					levelNo1 = 10
				ElseIf InStr(1, "Copy Cargo", trim(menuItem(0)), 1) Then
					levelNo1 = 25			
				End If
				
		Case "VOYAGE"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_Voyage").Click 307, 51, micRightBtn
				If InStr(1, "Deal", trim(menuItem(0)), 1) Then
					levelNo1 = 10
				End If
				
		Case "PORTLEG"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Click xright, xbtm, micRightBtn
				If InStr(1, "Cargo Parcels", trim(menuItem(0)), 1) Then
					levelNo1 = 18
				ElseIf InStr(1, "Events", trim(menuItem(0)), 1) Then
					levelNo1 = 11
				ElseIf InStr(1, "Bunker Liftings", trim(menuItem(0)), 1) Then
					levelNo1 = 14
				ElseIf InStr(1, "Mail", trim(menuItem(0)), 1) Then
					levelNo1 = 9
				ElseIf InStr(1, "Voyage Notes", trim(menuItem(0)), 1) Then
					levelNo1 = 6
				ElseIf InStr(1, "Deal", trim(menuItem(0)), 1) Then
					levelNo1 = 10
				End If
				
		Case "TRIPPORT"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_TripPortCalls").Click xright, xbtm, micRightBtn
				If InStr(1, "Add Terminal After", trim(menuItem(0)), 1) Then
					levelNo1 = 2
				ElseIf InStr(1, "Add Vessel", trim(menuItem(0)), 1) Then
					levelNo1 = 9		
				End If

		Case "TRIPVESSEL"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_Vessel").Click xright, xbtm, micRightBtn
				If InStr(1, "Cargo Parcels", trim(menuItem(0)), 1) Then
					levelNo1 = 3
				End If
				
		Case "DEMURRAGE"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbTreeView("PTV_Explorer").Click xright, xbtm, micRightBtn
'			Commented by sanghamitra
'				If InStr(1, "Claim Details...", trim(menuItem(0)), 1) Then
'					levelNo1 = 4
'				End If

			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "{c}"
			wait(3)
			
			WshShell.SendKeys "{ENTER}"
			
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Navigated To: " & strMenu)
			Exit function

		Case "DemurrageExplorerPayment"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_PaymentList").Click 175,56, micRightBtn
				Set WshShell =createobject("WScript.Shell")
				wait 2
				WshShell.SendKeys "{DOWN}"
				wait 2
				WshShell.SendKeys "{DOWN}"
				wait(2)
				WshShell.SendKeys "{ENTER}"
				Set WshShell=nothing
		
		Case "DemurrageExplorerInvoice"
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbDataWindow("PDw_invoice_master").Click 235,90, micRightBtn
'				If InStr(1, "Add Invoice Received...", trim(menuItem(0)), 1) Then
'					levelNo1 = 2
'				End If
				Set WshShell =createobject("WScript.Shell")
				wait 2
				WshShell.SendKeys "{DOWN}"
				wait 2
				WshShell.SendKeys "{DOWN}"
				wait(2)
				WshShell.SendKeys "{ENTER}"
				Set WshShell=nothing
	End Select
	
	If levelNo4 <> 0 Then
		itemPath = Window("Win_RADAR").WinMenu("Win_RightClickMenu").BuildMenuPath(levelNo1, levelNo2, levelNo3, levelNo4)
	ElseIf levelNo3 <> 0 Then
		itemPath = Window("Win_RADAR").WinMenu("Win_RightClickMenu").BuildMenuPath(levelNo1, levelNo2, levelNo3)
	ElseIf levelNo2 <> 0 Then
		itemPath = Window("Win_RADAR").WinMenu("Win_RightClickMenu").BuildMenuPath(levelNo1, levelNo2)
	Else
		itemPath = Window("Win_RADAR").WinMenu("Win_RightClickMenu").BuildMenuPath(levelNo1)
	End If
	
	If Window("Win_RADAR").Exist(0) Then
		wait(2)
		Window("Win_RADAR").WinMenu("Win_RightClickMenu").Select itemPath
		wait(15)
	
	If instr(strmenu, "PORTLEG")>0 or instr(strmenu, "TRIPPORT")>0 or instr(strmenu, "TRIPVESSEL")>0 or instr(strmenu, "DEMURRAGE")>0 Then
		strMenu = strObject(0)&","&strObject(1)
	End If
	
		If Err.Number > 0 Then
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Issue in navigation: " & Err.Description)
		Else
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Navigated To: " & strMenu)
		End If		
	Else 
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Issue in navigation: " & strMenu)	
	End If
	wait 1
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)	
End Function
'#####################################################################################################################
Public Function fnRAD_WaitForReady()
	CURRENT_FUNCTION = "fnRAD_WaitForReady"
	On Error Resume Next
	Call fnCOM_updateLogFile("I", CURRENT_FUNCTION, CURRENT_FUNCTION)
	
	wait 5
	For i=1 to 300
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton("PBtn_Ok").Exist(0) Then
			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton("PBtn_Ok").Click
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton("PBtn_OKBtn").Exist(0) Then
 			PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbButton("PBtn_OKBtn").Click
		ElseIf Dialog("PDia_Alert").WinButton("WinBtn_OK").Exist(0) Then
			Call fnRAD_ClickAlert("WinBtn_OK")
		ElseIf Dialog("PDia_Alert").WinButton("WinBtn_Yes").Exist(0) Then
			Call fnRAD_ClickAlert("WinBtn_Yes")
		End If
		
		If PbWindow("PWin_MainWindow").WinObject("win_Ready").Exist(1) Then
			wait 5
			Exit For
		Else
			wait 1
		End If
	Next
			
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function
'#####################################################################################################################
Public Function fnRAD_WaitUntilWinLoads()
	CURRENT_FUNCTION = "fnRAD_WaitUntilWinLoads"
	On Error Resume Next

wait 5
For i = 1 to 300
	If instr(Window("text:=R.A.D.A.R.*").GetROProperty("Text"), "Not Responding") Then
		wait 1
	Else
		Exit For
	End If
Next 
End Function

'#####################################################################################################################
Public Function fnRAD_GetDataValue(paramStr)
	'paramStr = objectName, [objectValue = xyz or <blank> or <rand> or ""]
	CURRENT_FUNCTION = "fnRAD_GetDataValue"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, paramStr)
	
	sFlag = False
	indvParam = Split(paramStr, ",")
	If UBound(indvParam) < 2 Then
		Call fnCOM_updateLogFile ("E", CURRENT_FUNCTION, paramStr)
	End If
	windowName = Trim(indvParam(0))
	rownNo=Trim(indvParam(1))
	objectName = Trim(indvParam(2))
	
	If Left(strObjValue, 2) = "c_" Then		
		colName = Replace(strObjValue, "c_", "")			
		outputUpdt = False		
		strObjValue = INPUT_QRY_REC.Fields.Item(colName)		
		If InStr(1, strObjValue, "<rand>", 1) Then outputUpdt = True
	End If		
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")
			sFlag = True
		Else
			resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
	End If	
	
	If PbWindow("PWin_MainWindow").Exist(0) Then
		sCellData = trim(rootObj.PbDataWindow(windowName).GetCellData(rownNo, objectName))
			If sCellData<>"" Then
				fnRAD_GetDataValue = sCellData
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks(objectName & " returned: " & sCellData)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks(objectName&" returned null")
			End IF
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function
'#####################################################################################################################
Public Function fnRAD_ClickonRow(paramStr)		'fnRAD_SelectRow uses instr method which was selecting any row which had the the expstr inside the output string. This function removed instr to check for exact values
	CURRENT_FUNCTION = "fnRAD_ClickonRow"
	On Error Resume Next
	Call fnCOM_updateLogFile("S", CURRENT_FUNCTION, paramStr)
	
	tStr = Split(paramStr, ",")
	rowStr = Trim(tStr(0))
	If UBound(tStr) > 0 Then
		tableName = Trim(tStr(1))
	Else
		tableName = "PDw_DataWindow"
	End If
	
	sFlag = False
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbWindow("PWin_SubWin2_1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").PbWindow("PWin_SubWin2_1")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")
			sFlag = True
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")
			sFlag = True
		Else
			resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
	End If		
	If sFlag Then
		sFlag = False
		totalRows = rootObj.PbDataWindow(tableName).RowCount
		totalCols = rootObj.PbDataWindow(tableName).ColumnCount
		For cCount = 1 To totalCols
			For rCount = 1 To totalRows
				rValue = rootObj.PbDataWindow(tableName).GetCellData(rCount, cCount)					
				If trim(rValue)=trim(rowStr) Then										
					rootObj.PbDataWindow(tableName).ActivateCell rCount, cCount
					Call fnCOM_updateLogFile("I", CURRENT_FUNCTION, rowStr & " found at " & rCount & "," & cCount)
					sFlag = True
					If rootObj.PbButton("PBtn_Choose").Exist(0) Then
						rootObj.PbButton("PBtn_Choose").Click
						wait 1
					End If
					Exit For
				End If
			Next			
			If sFlag Then 
				resultStr = "Pass"
				remarksStr = "Selected: '" & rowStr & "'"
				Call fnRAD_CloseAlerts()
				Exit For
			End If
		Next
		If Not sFlag Then
			resultStr = "Fail"
			remarksStr = "[ERR] '" & rowStr & "' not found"
		End If
	End If
	
	fnRAD_ClickonRow = sFlag
	Call fnCOM_SetTestResult(resultStr)
	Call fnCOM_SetTestRemarks(remarksStr)
	Call fnCOM_reportErr(CURRENT_FUNCTION)
End Function

'#####################################################################################################################
Public Function fnRAD_Treeview(strparam)
	CURRENT_FUNCTION = "fnRAD_Treeview"
	On Error Resume Next
	Call fnCOM_updateLogFile ("I", CURRENT_FUNCTION, strparam)
	
	fnRAD_Treeview=False
	
	arrItem=split(strparam, ",")
	sOpern=trim(arrItem(0))
	sFold=trim(arrItem(1))

	If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbTreeView("PTV_Explorer").Exist Then
		set ObjWin=PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").PbTreeView("PTV_Explorer")
	End If

	Select Case sOpern
		Case "SELCT"
		newSFoldname=left(sFold,1)
		PbWindow("PWin_MainWindow").PbWindow("w_explorer_dem").PbTreeView("tv_explorer").highlight
		
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys ("newSFoldname")
			ObjWin.Select sFold
			selectedtext=ObjWin.GetSelection

				If instr(sFold,selectedtext)>0 Then
					fnRAD_Treeview=True
					resultStr = "Pass"
					remarksStr = "Treeview selected: "&sFold
					Exit function
				Else
					resultStr = "Fail"
					remarksStr = "[ERR] Treeview Not selected: "&sFold	
					Exit function					
				End If
		Case "EXPAND"
			ObjWin.Expand sFold
			ObjWin.Select sFold
				If trim(ObjWin.GetSelection)=sFold Then
					fnRAD_Treeview=True
					resultStr = "Pass"
					remarksStr = "Treeview expanded: "&sFold
				Else
					resultStr = "Fail"
					remarksStr = "[ERR] Treeview Not expanded: "&sFold
				End If
	End Select

	Call fnCOM_SetTestResult(resultStr)
	Call fnCOM_SetTestRemarks(remarksStr)
	Call fnCOM_reportErr(CURRENT_FUNCTION)

End Function
'#####################################################################################################################

Public Function fnRAD_TypeDataValues1(paramStr,cellvalue)
	
	CURRENT_FUNCTION = "fnRAD_TypeDataValues"
	On Error Resume Next
	Call fnCOM_updateLogFile ("S", CURRENT_FUNCTION, paramStr)
	
	sFlag = False
	indvParam = Split(paramStr, ",")
	If UBound(indvParam) < 2 Then
		Call fnCOM_updateLogFile ("E", CURRENT_FUNCTION, paramStr)
	End If
	windowName = Trim(indvParam(0))
	objectName = Trim(indvParam(1))
	strObjValue = Trim(indvParam(2))
	tabvalue=Trim(indvParam(3))
	cellData=cellvalue
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin2")
			sFlag = True
		
		ElseIf PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1").Exist(0) Then
			Set rootObj = PbWindow("PWin_MainWindow").PbWindow("PWin_SubWin1")
			sFlag = True
		Else
			resultStr = "Fail"
			remarksStr = "[ERR] Unable to find the required window"
		End If
	End If
	If PbWindow("PWin_MainWindow").Exist(0) Then
		If strObjValue <> "" Then
			If rootObj.Exist(0) Then
					
				If tabvalue=1 Then
				rootObj.PbDataWindow(windowName).SelectCell cellData, objectName
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{TAB}"	
				wait(2)
				WshShell.SendKeys strObjValue
				wait(1)
				End If
				If tabvalue=2 Then
					rootObj.PbDataWindow(windowName).SelectCell cellData, objectName
				Set WshShell = CreateObject("WScript.Shell")
				WshShell.SendKeys "{TAB}"
				wait(2)
				WshShell.SendKeys "{TAB}"
				wait(2)
				WshShell.SendKeys strObjValue
'				rootObj.PbDataWindow(windowName).Type strObjValue	
				wait(1) 
				End If
				
				Call fnCOM_SetTestResult("Pass")
				Call fnCOM_SetTestRemarks(objectName & " set to " & strObjValue)
			Else
				Call fnCOM_SetTestResult("Fail")
				Call fnCOM_SetTestRemarks("[ERR] Object not found: " & objectName)
			End IF
		Else
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
		End If	
	Else
		Call fnCOM_SetTestResult("Fail")
		Call fnCOM_SetTestRemarks("[ERR] Session not available: " & objectName)
	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
	
	
	
	
	
	
	
	
End Function






Public Function ShellScript()
	Set WshShell =createobject("WScript.Shell")
				wait 2
				WshShell.SendKeys "%{ENTER}"
					wait 3
				Set WshShell = Nothing
End Function


Public function fnRAD_NavigateTo1(strMenu)

Window("Win_RADAR").Activate

Select Case strMenu
Case "Save"
			Set WshShell =createobject("WScript.Shell")
				wait 2
				WshShell.SendKeys "%{s}"
					wait 3
				Set WshShell = Nothing
wait(30)

Case "Find"
			Set WshShell =createobject("WScript.Shell")
				wait 2
				WshShell.SendKeys "%{f}"
					wait 3
				Set WshShell = Nothing
wait(30)

Case "Insert"
			Set WshShell =createobject("WScript.Shell")
				wait 2
				WshShell.SendKeys "%{i}"
					wait 3
				Set WshShell = Nothing


end select
On Error Resume Next
'	If Window("Win_RADAR").Exist(10) Then
'		On Error Resume Next
'		Window("Win_RADAR").WinMenu("Win_Menu").Select strMenu
		'wait(2)
		If Err.Number <> 0 Then
			Call fnCOM_SetTestResult("Fail")
			Call fnCOM_SetTestRemarks("[ERR] Unable to navigate to: " & strMenu)
		Else
			Call fnCOM_SetTestResult("Pass")
			Call fnCOM_SetTestRemarks("Navigated To: " & strMenu)
		End If
'	Else 
'		Call fnCOM_SetTestResult("Fail")
'		Call fnCOM_SetTestRemarks("[ERR] Unable to navigate to: " & strMenu)
'	End If
	
	Call fnCOM_reportErr(CURRENT_FUNCTION)
	

End  Function


