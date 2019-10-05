'###################################################################################################
'# SCRIPT NAME: sepa-iso2K22-BulkGen.vbs
'#
'# DESCRIPTION:
'# Free script utility for silent XML/XSD validation of large sized files.
'# The VBSX_Validator is designed to validate large XML files.The project 
'# exposes the power and flexibility of VB Script language and demonstrates how it 
'# could be utilized for some specific XML related operations and automation.
'# 
'# NOTES:
'# Dependency on MSXML6. Supports full multiple error parsing with offline log file output.
'# Also supports Batch (Multiple XML Files) Validation against a single specified XSD
'# The Parser does not resolve externals. It does not evaluate or resolve the schemaLocation 
'# or attributes specified in DocumentRoot. The parser validates strictly against the 
'# supplied XSD only without auto-resolving schemaLocation. The parser needs 
'# Namespace (targetNamespace) which is currently extracted from the supplied XSD.

'# PLATFORM: Win7/8/Server | PRE-REQ: Script/Admin Privilege
'# LAST UPDATED: Aug 2019 | AUTHOR: Tushar Sharma
'##################################################################################################



If WScript.Arguments.length = 0 Then
   Set objShell = CreateObject("Shell.Application")
   objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 3
      WScript.Quit
End If  


'#######################################################################################################

Dim LogHandle, strAppOutputDir

Const strInvalid = "invalid"
Const strFile = "file"
Const strFolder = "folder"
Const strFileExtXSD = ".xsd"
Const strFileExtXML = ".xml"



Call StartSEPABulkGen()


'#######################################################################################################


Sub StartSEPABulkGen()

Dim strModeSelected 
	
	ShowWelcomeBox()
	ShowMode ("nomode")
	strModeSelected = SelectMode()
	ShowMode (strOpsMode)
	
	If Not(IsObject(LogHandle)) Then
		Set LogHandle = CreateLogWriter()
	End If
	
	Select Case strModeSelected
		Case "1"
			Call ModeGenOnly()
		Case "2"
			Call ModeValOnly()
		Case "3"
			Call ModeValAndGen()
		Case Else
			ConsoleOutput "INVALID CHOICE!", "verbose"
			If IsReloadExit("") Then
				Call StartSEPABulkGen()
			Else
				ExitApp()
			End If
	End Select
	
		If IsReloadExit("") Then
			Call StartSEPABulkGen()
		Else
			ExitApp()
		End If

End Sub

'#######################################################################################################

Public Sub ModeValAndGen()
Dim strFilePath
Dim sFormatChoice

ShowWelcomeBox("valandgen")
ShowFileChoices()
sFormatChoice = ConsoleInput()

If CreateWorkingDir() Then
	Call ConsoleOutput ("PROVIDE FULL PATH TO TEMPLATE XML FILE (.xml) ? ", "verbose")
	strFilePath = ConsoleInput()
	Set MyXMLFile = LoadXML(strFilePath)
	
	Call ConsoleOutput ("PROVIDE FULL PATH TO THE SCHEMA FILE (.xsd) ? ", "verbose")
	Set MyXSDFile = LoadXSD(ConsoleInput(),GetNamespaceURI(MyXMLFile))
	
	ValidateXML MyXMLFile, MyXSDFile
	Set MyXSDFile = Nothing
	
	Call ConsoleOutput ("SPECIFY NUMBER OF TRANSACTIONS", "verbose")
	Call GenerateFile (sFormatChoice, MyXMLFile, ConsoleInput)
	Call SaveXML (MyXMLFile)
	
	Set MyXMLFile = Nothing
End If
	
End Sub	

'###########################################################################

Public Sub ModeValOnly()
Dim strFilePath
Dim sFormatChoice

ShowWelcomeBox("valonly")
ShowFileChoices()
sFormatChoice = ConsoleInput()

If CreateWorkingDir() Then
	Call ConsoleOutput ("PROVIDE FULL PATH TO TEMPLATE XML FILE (.xml) ? ", "verbose")
	strFilePath = ConsoleInput()
	Set MyXMLFile = LoadXML(strFilePath)
	
	Call ConsoleOutput ("PROVIDE FULL PATH TO THE SCHEMA FILE (.xsd) ? ", "verbose")
	Set MyXSDFile = LoadXSD(ConsoleInput(),GetNamespaceURI(MyXMLFile))
	
	ValidateXML MyXMLFile, MyXSDFile
	Set MyXSDFile = Nothing
	
	'Call ConsoleOutput ("SPECIFY NUMBER OF TRANSACTIONS", "verbose")
	'Call GenerateFile (sFormatChoice, MyXMLFile, ConsoleInput)
	'Call SaveXML (MyXMLFile)
	Set MyXMLFile = Nothing
End If

End Sub

'###########################################################################

Public Sub ModeGenOnly()

Dim MyXMLFile, iTrxCount, strCurFileName
Dim strFilePath, sFormatChoice

Set objFSO = CreateObject("Scripting.FileSystemObject")

ShowFileChoices()
sFormatChoice = ConsoleInput()

	ConsoleOutput "PROVIDE FULL PATH TO TEMPLATE (SEED) XML FILE (e.g. C:\PAIN008.xml) ?", "verbose", LogHandle
	strFilePath = ConsoleInput()
	
	If IsXMLXSD(strFilePath) = strFileExtXML Then
	
		Set MyXMLFile = LoadXML(strFilePath)
		strCurFileName = objFSO.GetFileName(strFilePath)
		strCurFileName = Left(strCurFileName,(Len(strCurFileName)-4))
		
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "NOTE: TOTAL TRANSACTIONS  = [No. Of 'PmtInf'] x [No. of 'DrctDbtTxInf'] ", "nolog", LogHandle		
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "SPECIFY NUMBER OF PAYMENT INSTRUCTION INFO BLOCKS (PmtInf) ?", "verbose", LogHandle
		iPmtInfCount = ConsoleInput()
		ConsoleOutput "SPECIFY NUMBER OF DEBIT TRANSACTIONS (DrctDbtTxInf) ?", "verbose", LogHandle		
		iTrxCount = ConsoleInput()
					
		If IsNumeric(iPmtInfCount) And IsNumeric(iTrxCount) Then
			Call GenerateFile (sFormatChoice, MyXMLFile, iPmtInfCount, iTrxCount)
			Call GetOutputFolder()
			Call SaveXML (MyXMLFile, strCurFileName)
		Else
			ConsoleOutput "", "verbose", LogHandle
			ConsoleOutput "<ERROR> INVALID INPUT! PLEASE TRY AGAIN ...", "verbose", LogHandle
			If IsReloadExit("") Then
				Call StartSEPABulkGen()
			Else
				ExitApp()
			End If
		
		End If
	Else
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "<ERROR> INVALID FILE OR PATH! PLEASE TRY AGAIN ...", "verbose", LogHandle
		If IsReloadExit("") Then
			Call StartSEPABulkGen()
		Else
			ExitApp()
		End If
	
	End If 

	Set MyXMLFile = Nothing
	
End Sub

'###########################################################################

Public Function GenerateFile (sFormatChoice, ObjSeedFile, iNumPmtBlocks, strNumTrx)

Select Case sFormatChoice
	Case "1"
		Call GeneratePAIN008 (ObjSeedFile, iNumPmtBlocks, strNumTrx)
	Case "2"
		Call GeneratePACS008EPC (ObjSeedFile, strNumTrx)
	Case "3"
		Call GeneratePACS003 (ObjSeedFile, strNumTrx)
	Case "4"
		Call GeneratePACS003EPC (ObjSeedFile, strNumTrx)
	Case "5"
		Call GeneratePACS003EBA (ObjSeedFile, strNumTrx)
	Case "6"
		Call GeneratePACS003EPC (ObjSeedFile, strNumTrx)
	Case Else
		ConsoleOutput "INVALID CHOICE!", "verbose", LogHandle
End Select

End Function

'###########################################################################

Public Function LoadXML(strXmlPath)

Dim ObjParseErr
Dim ObjXML

Set ObjXML = CreateObject ("MSXML2.DOMDocument.6.0")
	With ObjXML
		'Set First Level DOM Properties
		.async = False
		.validateOnParse = False
		.resolveExternals = False
	End With
	
	ConsoleOutput "", "verbose", LogHandle
	ConsoleOutput "<INFO> Loading XML with First-Level XMLDOM Properties", "verbose", LogHandle
	ObjXML.Load (strXmlPath)
	
	If ObjXML.ParseError.errorCode <> 0 Then
		Call ParseLoadErrors (ObjXML.parseError)
		If IsReloadExit("") Then
			Call StartSEPABulkGen()
		Else
			ExitApp()
		End If
	Else
		ConsoleOutput "<INFO> File Loaded Successfully ..." & strXmlPath, "verbose", LogHandle
		ConsoleOutput "<INFO> Setting Up XML Namespace Property ...", "verbose", LogHandle
		ObjXML.setProperty "SelectionNamespaces", "xmlns:ns='" + ObjXML.documentElement.namespaceURI + "'"
		ConsoleOutput "<INFO> Setting Up XMl Selection Language Propoerty ... XPath", "verbose", LogHandle
		ObjXML.setProperty "SelectionLanguage", "XPath"
		ConsoleOutput "<INFO> First Level XML DOM Properties configured successfully. ", "verbose", LogHandle
		Set LoadXML = ObjXML
	End If

End Function

'###########################################################################

Public Function LoadAndValidate(strXmlPath)
Dim ObjXML

Set ObjXML = CreateObject ("MSXML2.DOMDocument.6.0")
	With ObjXML
		'Set First Level DOM Properties
		.async = False
		.validateOnParse = True
		.resolveExternals = True
	End With
	ObjXML.Load (strXmlPath)
	
	'Return Parse Error Object
	If ObjXML.ParseError.errorCode <> 0 Then
		ParseError (ObjXML.ParseError)
		Set LoadAndValidate = Nothing
	Else
		ConsoleOutput "File Loaded Successfully ..." & strXmlPath, "verbose"
		ConsoleOutput "Setting Up XML Namespace Property ...", "verbose"
		ObjXML.setProperty "SelectionNamespaces", "xmlns:ns='" + ObjXML.documentElement.namespaceURI + "'"
		ConsoleOutput "Setting Up XMl Selection Language Propoerty ... XPath", "verbose"
		ObjXML.setProperty "SelectionLanguage", "XPath"
		ConsoleOutput "First Level XML DOM Properties configured successfully. ", "verbose"
		Set LoadAndValidate = ObjXML
	End If

End Function

'###########################################################################

Public Function LoadXSD (strXSDPath, strNsURI)

Dim ObjXSD
	Set ObjXSD = CreateObject("MSXML2.XMLSchemaCache.6.0")
	ObjXSD.validateOnload = False
	
	'Load XSD from the Path
	ObjXSD.Add strNsURI, strXSDPath
	ConsoleOutput "XML Schema File Loaded Successfully ... " & strXSDPath, "verbose"
	Set LoadXSD = ObjXSD
	
End Function
	
'###########################################################################

Public Function GetNamespaceURI (ObjXML)

Dim strNsURI
If Not IsObject(ObjXML) Then
	Set ObjXMLDoc = LoadXML(ObjXML)
	strNsURI = ObjXMLDoc.documentElement.namespaceURI
Else
	strNsURI = ObjXML.documentElement.namespaceURI
End If
GetNamespaceURI = strNsURI
Set ObjXMLDoc = Nothing

End Function

'###########################################################################

Public Function ValidateXML (ObjXMLDoc, ObjXSDDoc)

Set ObjXMLDoc.Schemas = ObjXSDDoc
If ObjXMLDoc.readystate = 4 Then
	Set ObjXParseErr = ObjXMLDoc.validate()
	ParseError (ObjXParseErr)
End If

End Function

'###########################################################################

Public Function ParseError (ByVal ObjParseErr)

Dim strResult
Select Case ObjParseErr.errorCode
	Case 0
		strResult = "XML SCHEMA VALIDATION: SUCCESS! " & strFileName & vbCr
		ParseError = True
	Case Else
		strResult = vbCrLf & "ERROR! VALIDATION FAILED " & _
		vbCrLf & ObjParseErr.reason & vbCr & _
		"Error Code: " & ObjParseErr.errorCode & ", Line: " & _
						 ObjParseErr.Line & ", Character: " & _		
						 ObjParseErr.linepos & ", Source: " & _
						 Chr(34) & ObjParseErr.srcText & _
						 Chr(34) & " - " & Now & vbCrLf
		ParseError = False
End Select
ConsoleOutput strResult, "verbose"

End Function

'###########################################################################

Public Function GetValue (ObjXMLDoc, XPathQuery)

'On Error GOTO GetValue Error Handler
ObjXMLDoc.setProperty "SelectionNamespaces", "xmlns:ns='" + ObjXMLDoc.document.Element.namespaceURI + "'"
Set NodeList = ObjXMLDoc.selectNodes(XPathQuery)

If NodeList.Length > 0 Then
	GetValue = NodeList.Item(0).Text
End If
'GetValue Error Handler Here ...

End Function

'###########################################################################

Public Function ConsoleInput()
Dim strIn

strIn = WScript.StdIn.ReadLine

If (Right(strIn,1) = Chr(34)) And (Left(strIn,1) = Chr(34)) Then
	strIn = Replace(strIn,Chr(34),"")
End If

ConsoleInput = strIn

End Function

'###########################################################################

Public Sub ConsoleOutput (strMsg, strMode, objFSOHandle)

Select Case strMode
	Case LCase("logonly")
		objFSOHandle.WriteLine (strMsg)
	Case LCase("nolog")
		WScript.StdOut.WriteLine (strMsg)
	Case LCase ("verbose")
		WScript.StdOut.WriteLine (strMsg)
		objFSOHandle.WriteLine (strMsg)
End Select

End Sub

'###########################################################################

Public Function CreateLogWriter()

sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
strFileName = "sepa-2K22-BulkGen" & "_" & Day(Date) & MonthName(Month(Date),True) & Right((Year(Date)),2) & ".txt"

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjTextFile = ObjFSO.OpenTextFile(sCurrPath & strFileName, 8, True)
strLogPath = sCurrPath & strFileName

Set CreateLogWriter = ObjTextFile 

Set ObjTextFile = Nothing
Set ObjFSO = Nothing

End Function


'###########################################################################

Public Function GeneratePAIN008 (ByRef ObjXMLDoc, iNumPmtBlocks, PmtCounter)

Dim iCount
Set ObjFSO = CreateObject("Scripting.FileSystemObject")

'Set strHdrNodes = ObjXMLDoc.selectsingleNode("//ns:MsgId")
Set strHdrNodes = GetSingleNode(ObjXMLDoc,"//ns:MsgId")
strHdrNodes.Text = "MsgID" & GetRandomChars(ObjFSO)
Set strHdrNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,"//ns:NbOfTxs")
strHdrNodes.Text = CLng(strHdrNodes.Text) + CLng(PmtCounter)
Set strHdrNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,"//ns:CtrlSum")
strHdrNodes.Text = CCur(strHdrNodes.Text) + CCur(strHdrNodes.Text*PmtCounter)
Set strHdrNodes = Nothing

'PMNTINFO NODE CLONED HERE ...
Set sPmtInfoCpy1 = GetSingleNode(ObjXMLDoc,"//ns:PmtInf").CloneNode(True)
Set stemps = ObjXMLDoc.selectNodes("//ns:PmtInf")
stemps.RemoveAll
Set stemps = Nothing

For iNum = 1 To iNumPmtBlocks
	
	Call ConsoleOutput("=========" & " STARTED GENERATING PAYMENT INSTRUCTION BLOCK [PmtInf] " & "=========", "verbose", LogHandle)
	Set strPmtInfo1 = sPmtInfoCpy1.CloneNode(True)
	
	''CLONED PMNTINFO EDITED HERE ...
	Set strHdrNodes = GetSingleNode(strPmtInfo1,"//ns:PmtInfId")
	strHdrNodes.Text = "PmtID" & GetRandomChars(ObjFSO)
	''CHANGE CRDTR BIC/IBAN TAGS HERE ...
	Set strHdrNodes = Nothing
	
	''DDTRX NODE CLONED HERE ...
	Set sDDTrxCpy1 = GetSingleNode(strPmtInfo1,"//ns:DrctDbtTxInf").CloneNode(True)
	Set stemps = strPmtInfo1.selectNodes("//ns:DrctDbtTxInf")
	stemps.RemoveAll
	Set stemps = Nothing
	
	For iCount = 1 To PmtCounter
		If iCount = 1 Then
			Call ConsoleOutput(" ========" & " STARTED GENERATING DEBIT TRANSACTIONS [DrctDbtTxInf] " & "========", "verbose", LogHandle)
			Call ConsoleOutput("Started Generating " & PmtCounter & " Debit Transactions at " & Now, "verbose", LogHandle)
		End If

		Set sTempNode = sDDTrxCpy1.CloneNode(True)
		Set sEndToEndID = GetSingleNode(sTempNode,"//ns:EndToEndId")
			sEndToEndID.Text = iCount & GetRandomChars(ObjFSO)
		Set sEndToEndID = Nothing
		
		strPmtInfo1.AppendChild sTempNode
		Set sTempNode = Nothing
		
		Call ConsoleOutput ("Generating Debit Transaction ... " & iCount, "nolog", LogHandle)
		
		If iCount = CLng(PmtCounter) Then
			Call ConsoleOutput ("All " & PmtCounter & " Payments Generated Successfully at " & Now, "verbose", LogHandle)
			Call ConsoleOutput ("=============" & " COMPLETED TRANSACTION GENERATION [DrctDbtTxInf]" & "============", "verbose", LogHandle)
		End If
	Next
	
	'PMNTINFO NODE APPENDED TO THE DOC FRAG HERE ...
	Set oDocFrag = ObjXMLDoc.CreateDocumentFragment
	oDocFrag.AppendChild strPmtInfo1
	
	Call ConsoleOutput("========" & " COMPLETED GENERATING PAYMENT INSTRUCTION BLOCK [PmtInf] " & "========" & vbCrLf & vbCrLf & vbCrLf, "verbose", LogHandle)

	Set strPmtInfo1 = Nothing
	Set sDDTrxCpy1 = Nothing
	
	GetSingleNode(ObjXMLDoc,"//ns:CstmrDrctDbtInitn").AppendChild oDocFrag
	Set oDocFrag = Nothing

Next

Set ObjFSO = Nothing

End Function	

'###########################################################################

Public Function GeneratePACS008EPC (ByRef ObjXMLDoc, PmtCounter)
Dim iCount

Set strHdrNodes = ObjXMLDoc.selectsingleNode("//ns:MsgId")
strHdrNodes.Text = "MsgID" & GetRandomChars()
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDoc.selectsingleNode("//ns:NbOfTxs")
strHdrNodes.Text = CLng(strHdrNodes.Text) + CLng(PmtCounter)
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDoc.selectsingleNode("//ns:TtlIntrBkSttlmAmt")
strHdrNodes.Text = CCur(strHdrNodes.Text) + CCur(strHdrNodes.Text*PmtCounter)
Set strHdrNodes = Nothing

Set sSelectedNode = ObjXMLDoc.selectsingleNode("//ns:CdtTrfTxInf").CloneNode(True)
Set oDocFrag = ObjXMLDoc.CreateDocumentFragment

For iCount = 1 To PmtCounter

	If iCount = 1 Then
		Call ConsoleOutput("========" & " XML FILE GENERATION PROCESS STARTED " & "========", "verbose")
		Call ConsoleOutput("Started Generating " & PmtCounter & " Payments at " & Now & "...", "verbose")
	End If
	strE2EIds = "RNDID" & iCount
	
	Set sTempNode = sSelectedNode.CloneNode(True)
	Set sEndToEndID = sTempNode.selectsingleNode("//ns:InstrId")
		sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	Set sEndToEndID = sTempNode.selectsingleNode("//ns:EndToEndId")
		sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	Set sEndToEndID = sTempNode.selectsingleNode("//ns:TxId")
		sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	oDocFrag.AppendChild sTempNode
	sTempNode = Nothing

	Call ConsoleOutput ("Generating Payment ... " & iCount, "nolog")
	
	if iCount = CLng(PmtCounter) Then
		Call ConsoleOutput ("All " & PmtConter & " Payments generated successfully at " & Now, "verbose")
		Call ConsoleOutput ("=================" & " PROCESS ENDED " & "=================" & vbCrLf & vbCrLf, "verbose")
	End If

Next

Set sSelectedNode = Nothing
ObjXMLDoc.selectsingleNode("//ns:FIToFICstmrCdtTrf").AppendChild oDocFrag
Set oDocFrag = Nothing

End Function


'###########################################################################

Public Function GeneratePACS003 (ByRef ObjXMLDoc, PmtCounter)
Dim iCount
Dim strHdrNodes

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns:MsgId")
strHdrNodes.Text = "MsgID" & GetRandomChars()
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns:NbOfTxs")
strHdrNodes.Text = CInt(strHdrNodes.Text) + CLng(PmtCounter)
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns:TtIntrBkSttlmAmt")
strHdrNodes.Text = CCur(strHdrNodes.Text) + CCur(strHdrNodes.Text*PmtCounter)
Set strHdrNodes = Nothing

Set sSelectedNode = ObjXMLDoc.selectsingleNode("//ns:DrctDbtTxInf").CloneNode(True)
Set oDocFrag = ObjXMLDoc.CreateDocumentFragment

For iCount = 1 To PmtCounter

	If iCount = 1 Then
		Call ConsoleOutput("===========" & " XML FILE GENERATION PROCESS STARTED " & "===========", "verbose")
		Call ConsoleOutput("Started Generating " & PmtCounter & " Payments at " & Now & "...", "verbose")
	End If
	strE2EIds = "RNDID" & iCount
	
	Set sTempNode = sSelectedNode.CloneNode(True)
	
	Set sEndToEndID = sTempNode.selectsingleNode("//ns:InstrId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing

	Set sEndToEndID = sTempNode.selectsingleNode("//ns:EndToEndId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing

	Set sEndToEndID = sTempNode.selectsingleNode("//ns:TxId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	oDocFrag.AppendChild sTempNode
	Set sTempNode = Nothing

	Call ConsoleOutput("Generating Payment ... " & iCount, "nolog")

	if iCount = CLng(PmtCounter) Then
		Call ConsoleOutput ("All " & PmtConter & " Payments generated successfully at " & Now, "verbose")
		Call ConsoleOutput ("=================" & " PROCESS ENDED " & "=================" & vbCrLf & vbCrLf, "verbose")
	End If

Next

Set sSelectedNode = Nothing
ObjXMLDoc.selectsingleNode("//ns:FIToFICstmrDrctDbt").AppendChild oDocFrag
Set oDocFrag = Nothing

End Function

'###########################################################################

Public Function GeneratePACS003EBA (ByRef ObjXMLDoc, PmtCounter)
Dim iCount
Dim strHdrNodes

ObjXMLDoc.setProperty "SelectionNamespaces", "xmlns:ns1='urn:iso:std:iso:2022:tech:xsd:pacs.003.001.02' xmlns:ns2='urn:S2SDDDnf:xsd:$MPEDDDnfBlkDirDeb'"

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns1:MsgId")
strHdrNodes.Text = "MsgID" & GetRandomChars()
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns1:NbOfTxs")
strHdrNodes.Text = CInt(strHdrNodes.Text) + CLng(PmtCounter)
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns1:TtIntrBkSttlmAmt")
strHdrNodes.Text = CCur(strHdrNodes.Text) + CCur(strHdrNodes.Text*PmtCounter)
Set strHdrNodes = Nothing

Set sSelectedNode = ObjXMLDoc.selectsingleNode("//ns1:DrctDbtTxInf").CloneNode(True)
Set oDocFrag = ObjXMLDoc.CreateDocumentFragment

For iCount = 1 To PmtCounter

	If iCount = 1 Then
		Call ConsoleOutput("===========" & " XML FILE GENERATION PROCESS STARTED " & "===========", "verbose")
		Call ConsoleOutput("Started Generating " & PmtCounter & " Payments at " & Now & "...", "verbose")
	End If
	strE2EIds = "RNDID" & iCount
	
	Set sTempNode = sSelectedNode.CloneNode(True)
	
	Set sEndToEndID = sTempNode.selectsingleNode("//ns1:EndToEndId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing

	Set sEndToEndID = sTempNode.selectsingleNode("//ns1:TxId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	oDocFrag.AppendChild sTempNode
	Set sTempNode = Nothing

	Call ConsoleOutput("Generating Payment ... " & iCount, "nolog")

	if iCount = CLng(PmtCounter) Then
		Call ConsoleOutput ("All " & PmtConter & " Payments generated successfully at " & Now, "verbose")
		Call ConsoleOutput ("=================" & " PROCESS ENDED " & "=================" & vbCrLf & vbCrLf, "verbose")
	End If

Next

Set sSelectedNode = Nothing
ObjXMLDoc.selectsingleNode("//ns2:FIToFICstmrDrctDbt").AppendChild oDocFrag
Set oDocFrag = Nothing

End Function

'###########################################################################

Public Function GeneratePACS003EPC (ByRef ObjXMLDoc, PmtCounter)
Dim iCount
Dim strHdrNodes

'Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns:MsgId")
Set strHdrNodes = GetSingleNode(ObjXMLDoc, "//ns:MsgId")
If Not(IsNull(strHdrNodes)) Then
MsgBox "hi"
End If

strHdrNodes.Text = "MsgID" & GetRandomChars()
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns:NbOfTxs")
strHdrNodes.Text = CInt(strHdrNodes.Text) + CLng(PmtCounter)
Set strHdrNodes = Nothing

Set strHdrNodes = ObjXMLDOc.selectsingleNode("//ns:TtIntrBkSttlmAmt")
strHdrNodes.Text = CCur(strHdrNodes.Text) + CCur(strHdrNodes.Text*PmtCounter)
Set strHdrNodes = Nothing

Set sSelectedNode = ObjXMLDoc.selectsingleNode("//ns:DrctDbtTxInf").CloneNode(True)
Set oDocFrag = ObjXMLDoc.CreateDocumentFragment

For iCount = 1 To PmtCounter

	If iCount = 1 Then
		Call ConsoleOutput("===========" & " XML FILE GENERATION PROCESS STARTED " & "===========", "verbose")
		Call ConsoleOutput("Started Generating " & PmtCounter & " Payments at " & Now & "...", "verbose")
	End If
	strE2EIds = "RNDID" & iCount
	
	Set sTempNode = sSelectedNode.CloneNode(True)

	Set sEndToEndID = sTempNode.selectsingleNode("//ns:InstrId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	Set sEndToEndID = sTempNode.selectsingleNode("//ns:EndToEndId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing

	Set sEndToEndID = sTempNode.selectsingleNode("//ns:TxId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	oDocFrag.AppendChild sTempNode
	Set sTempNode = Nothing

	Call ConsoleOutput("Generating Payment ... " & iCount, "nolog")

	if iCount = CLng(PmtCounter) Then
		Call ConsoleOutput ("All " & PmtConter & " Payments generated successfully at " & Now, "verbose")
		Call ConsoleOutput ("=================" & " PROCESS ENDED " & "=================" & vbCrLf & vbCrLf, "verbose")
	End If

Next

Set sSelectedNode = Nothing
ObjXMLDoc.selectsingleNode("//ns:FIToFICstmrDrctDbt").AppendChild oDocFrag
Set oDocFrag = Nothing

End Function

'###########################################################################

Public Function GetSingleNode (ObjXMLDocFrag, strXPathString)

Dim ObjTempNode 
Set ObjTempNode = ObjXMLDocFrag.selectSingleNode(strXPathString)

If Not(ObjTempNode is Nothing) Then
	Set GetSingleNode = ObjTempNode
Else
	Call ConsoleOutput ("<ERROR> INVALID XML. NODE NOT FOUND ! : " & strXPathString, "verbose", LogHandle)
	If IsReloadExit("") Then
		Call StartSEPABulkGen()
	Else
		ExitApp()
	End If
End If

End Function

'###########################################################################

Public Function GetRandomChars (ObjFSO)
upperlimit = 50000
lowerlimit = 1

Randomize
RndChrs = Int((upperlimit - lowerlimit + 1) * Rnd() + lowerlimit)
TmpName = Trim(Mid(ObjFSO.GetTempName, 4, 5))
strTmstp = Trim(Replace(Left(Time, 8), ":", ""))

GetRandomChars = (RndChrs & TmpName & strTmstp)

End Function

'###########################################################################

Function GetOutputFolder ()

sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
strMainFolderName = "sepa_BulkGen_Output"

Set ObjFSO = CreateObject("Scripting.FileSystemObject")

If Not (ObjFSO.FolderExists(sCurrPath & strMainFolderName)) Then
	ConsoleOutput "<INFO> Creating Output Folder ... ", "nolog", LogHandle
	Set ObjOutputDir = ObjFSO.CreateFolder(sCurrPath & strMainFolderName)
Else
	Set ObjOutputDir = ObjFSO.GetFolder(sCurrPath & strMainFolderName)
	ConsoleOutput "<INFO> Output Folder Located... ", "nolog", LogHandle
End if

strAppOutputDir = ObjOutputDir.Path
Set GetOutputFolders = ObjOutputDir

End Function

'###########################################################################


Public Function SaveXML (ByRef ObjXMLDoc, strFileTypeName)

Dim strFullFilePath, strFullFileName

	strFullFileName = CStr(strFileTypeName & "-" & Replace(Time, ":", "") & ".xml")
	strFullFilePath = strAppOutputDir & "\" & strFullFileName
	
	ConsoleOutput "<INFO> Saving the Generated XML File ... " & strFullFileName, "verbose", LogHandle
	ObjXMLDoc.Save (strFullFilePath)
	ConsoleOutput "<INFO> DONE! : File Saved At " & strFullFilePath, "verbose", LogHandle

End Function

'###########################################################################

Function IsXMLXSD(strFilePath)

Dim objFSO, strFileExt

Set objFSO = CreateObject("Scripting.FileSystemObject") 

If IsFolderFile(strFilePath) = strFile Then
	strFileExt = objFSO.GetFileName(strFilePath)
	
	Select Case Right(strFileExt,4)
		Case ".xml"
			IsXMLXSD = strFileExtXML	
		Case ".xsd"
			IsXMLXSD = strFileExtXSD
		Case Else
			IsXMLXSD = strInvalid
	End Select

Else
	IsXMLXSD = strInvalid
End If

End Function


'###########################################################################

Function IsFolderFile(strPathInput)

Set objFSO = CreateObject("Scripting.FileSystemObject") 

If objFSO.FileExists(strPathInput) Then 
    IsFolderFile = strFile
ElseIf objFSO.FolderExists(strPathInput) Then
	IsFolderFile = strFolder
Else 
	IsFolderFile = strInvalid
End If

End Function

'###########################################################################


Function IsReloadExit (ObjXML)
IsWait = True

If IsObject(ObjXML) Then
	Do While Not (ObjXML.readystate = 4)
		ConsoleOutput "Working on large size document, do you wish to continue (y/n)?", "nolog", LogHandle
		strResponse = UCase(ConsoleInput())
		If (strResponse = "N") Or (strResponse = "NO") Then
			IsWait = False
			Exit Do
		Else 
			WScript.Sleep(5000)
		End If
	Loop 
End If

ConsoleOutput "", "nolog", LogHandle
ConsoleOutput "RE-LOAD THE PROGRAM OR EXIT (y=Reload / n=Exit) ?", "nolog", LogHandle
strResponse = UCase(ConsoleInput())

If ValidateInput(strResponse) Then
	Select Case strResponse
	    Case "Y"
	    	IsReloadExit = True
	    Case "N"
	    	IsReloadExit = False
	End Select
Else
	ConsoleOutput "INVALID CHOICE!", "verbose", LogHandle
End If

If Not(IsWait) Then
	Call ExitApp()
End If

End Function

'###########################################################################

Function ValidateInput (strArgsIn)

Dim strValidInput, strArg, strFound
strFound = False
strValidNumIn = Array("1","2")
strValidStrIn = Array("Y","N","YES","NO")

If IsNumeric(strArgsIn) Then
	For Each strArg In strValidNumIn
		If (StrComp(strArg, strArgsIn) = 0) Then
			strFound = True
			Exit For
		End If
	Next
Else
	For Each strArg In strValidStrIn
		If (StrComp(UCase(strArg), strArgsIn) = 0) Then
			strFound = True
			Exit For
		End If
	Next
End If
	
	
	If Not(strFound) Then
		ValidateInput = False
	Else 
		ValidateInput = True
	End If

End Function 

'###########################################################################

Sub ExitApp()
	 WScript.StdOut.WriteBlankLines(1)
	 WScript.StdOut.WriteLine "Press 'Enter' key to exit ..."
	 ConsoleInput()
	 WScript.Quit
End Sub

'###########################################################################


Public Sub ShowWelcomeBox()

WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "      " & "****************************************************************"
WScript.StdOut.WriteLine "      " & "----------------------------------------------------------------"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & vbTab & "   " & "sepa-Iso2K22-BulkGen version 1.0.2"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "    " & "SEPA Compliant Bulk XML File Generator and Validator"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "     " & "Platform: Win7/8 | Pre-Req: Script/Admin Privilege"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine VBTab & "   " & "Updated: Sept 2019 | Tushar Sharma | www.testoxide.com"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "      " & "****************************************************************"
WScript.StdOut.WriteLine "      " & "----------------------------------------------------------------"
WScript.StdOut.WriteBlankLines(2)

End Sub

'###########################################################################

'This Function sets input values for operating modes 
Sub ShowMode (strMode)

WScript.StdOut.WriteBlankLines(1)

Select Case strMode
	Case LCase("valandgen")
		WScript.StdOut.WriteLine "MODE:- (VAL & GEN)"
	Case LCase("genonly")
		WScript.StdOut.WriteLine "MODE:- (GEN ONLY)"
	Case LCase("valonly")
		WScript.StdOut.WriteLine "MODE:- (VAL ONLY)"
	Case LCase("nomode")
		WScript.StdOut.WriteLine "MODE:- (NOT SET!)"
End Select

WScript.StdOut.WriteBlankLines(1)

End Sub

'###########################################################################

Public Function SelectMode()

WScript.StdOut.WriteLine "SELECT OPERATING MODE? [Example: Input 1 for Generation Only]"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "1. SEPA File Generation Only"
WScript.StdOut.WriteLine "2. SEPA File Validation (XSD) Only"
WScript.StdOut.WriteLine "3. SEPA File Generation and Validation (XSD)"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "Tip: Type a bullet number from above and hit Enter."
WScript.StdOut.WriteBlankLines(1)

strMode = ConsoleInput()
SelectMode = strMode 

End Function

'###########################################################################

Public Sub ShowReadMe()

Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set ObjTextFile = ObjFSO.OpenTextFile(GetCurrentDir() & "\ReadMe.txt", 1, False)
Set ObjTextFile = Nothing
Set ObjFSO = Nothing

End Sub

'###########################################################################

Public Sub ShowFileChoices()

WScript.StdOut.WriteLine "SELECT THE FILE FORMAT BELOW [Example: Input 1 for PAIN008 Format]"
WScript.StdOut.WriteBlankLines(1)
WScript.StdOut.WriteLine "1. DD PAYMENT INITIATION PAIN.008"
WScript.StdOut.WriteBlankLines(1)

End Sub

'###########################################################################

Public Function GetCurrentDir(strPath)

'GetCurrentDir = Left(strPath,InStrRev(strPath,"\"))
sCurrPath = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName)) - (Len(WScript.ScriptName)))
GetCurrentDir = sCurrPath

'The Other Methods -
'Dim sCurrPath
'sCurrPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
'Msgbox sCurrPath

End Function

'###########################################################################

Public Function ShowXML()

'Display XML on the console using the XMLDoc.xml property

End Function

'###########################################################################

Public Function SetValue(ObjXML, XPathQuery, strValue)

End Function

'###########################################################################


























