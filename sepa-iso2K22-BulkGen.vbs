'###################################################################################################
'# SCRIPT NAME: sepa-iso2K22-BulkGen.vbs
'#
'# DESCRIPTION:
'# The 'sepa-iso2K22-BulkGen' is a script based desktop utility capable of generating SEPA compliant
'# 'Bulk' (large sized) XML 'payment messages' (ISO2022) for 'Performance Testing' (NFT). It makes 
'# use of the windows based, legacy, but powerful MSXML API (msxml*.dll). The app also supports 
'# XSD validation of SEPA payment message XMLs against the user specified XSDs for PAIN and PACS formats.
'# It is specifically targetted for generating 'large-sized XMLs' from performance testing perspective 
'# and not necessarly built for functional or business level file generation / validations.
'
'# NOTES:
'# Dependency on MSXML6. Supports full multiple error parsing with offline log file output.
'# The Parser does not resolve externals. It does not evaluate or resolve the schemaLocation 
'# or attributes specified in DocumentRoot. The parser validates strictly against the supplied
'# XSD only without auto-resolving schemaLocation. The parser needs Namespace (targetNamespace) 
'# which is currently extracted from the supplied XSD.

'# SUPPORTED MESSAGE FORMATS:
'# FIToFICustomerCreditTransferV04	pacs.008.001.04
'# FIToFICustomerDirectDebitV04	pacs.003.001.04
'# CustomerDirectDebitInitiationV04	pain.008.001.04
'# CustomerCreditTransferInitiationV05	pain.001.001.05

'# PLATFORM: Win7/8/Server | PRE-REQ: Script/Admin Privilege | License: Apache 2.0
'# LAST UPDATED: Aug 2019 | AUTHOR: Tushar Sharma
'##################################################################################################



If WScript.Arguments.length = 0 Then
   Set objShell = CreateObject("Shell.Application")
   objShell.ShellExecute "cscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 3
      WScript.Quit
End If  


'#######################################################################################################

Dim LogHandle, strAppOutputDir, arrPAIN001AmtOption
arrPAIN001AmtOption = Array("//ns:InstdAmt","//ns:EqvtAmt/ns:Amt")

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

Dim MyXMLFile, arrTrxFileCnt
Dim strFilePath, sFormatChoice, strCurFileName

Set objFSO = CreateObject("Scripting.FileSystemObject")

ShowFileChoices()
sFormatChoice = ConsoleInput()

	ConsoleOutput "", "verbose", LogHandle	
	ConsoleOutput "PROVIDE FULL PATH TO TEMPLATE (SEED) XML FILE (e.g. C:\PAIN008.xml) ?", "verbose", LogHandle
	strFilePath = ConsoleInput()
	
	If IsXMLXSD(strFilePath) = strFileExtXML Then
	
		Set MyXMLFile = LoadXML(strFilePath)
		strCurFileName = objFSO.GetFileName(strFilePath)
		strCurFileName = Left(strCurFileName,(Len(strCurFileName)-4))
		arrTrxFileCnt = GetTrxFileCount (sFormatChoice)

		If isArray(arrTrxFileCnt) Then
			Call GenerateFile (sFormatChoice, MyXMLFile, arrTrxFileCnt(0), arrTrxFileCnt(1))
			Call GetOutputFolder()
			Call SaveXML (MyXMLFile, strCurFileName)
		Else
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
		Call GeneratePACS003 (ObjSeedFile, iNumPmtBlocks, strNumTrx)
	Case "3"
		Call GeneratePACS008 (ObjSeedFile, iNumPmtBlocks, strNumTrx)
	Case "4"
		Call GeneratePAIN001 (ObjSeedFile, iNumPmtBlocks, strNumTrx)
	Case Else
		ConsoleOutput "INVALID CHOICE!", "verbose", LogHandle
End Select

End Function

'###########################################################################

Public Function GetTrxFileCount (sFormatChoice)

Dim iPmtInfCount, iTrxCount

Select Case sFormatChoice
	Case "1"	'Pain008
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
			GetTrxFileCount = Array(iPmtInfCount,iTrxCount)
		Else
			ConsoleOutput "", "verbose", LogHandle
			ConsoleOutput "<ERROR> INVALID INPUT! PLEASE TRY AGAIN ...", "verbose", LogHandle
			GetTrxFileCount = False
		End If
	Case "2"	'PACS003
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "NOTE: TOTAL TRANSACTIONS  = [No. of 'DrctDbtTxInf'] ", "nolog", LogHandle		
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "SPECIFY NUMBER OF DEBIT TRANSACTIONS (DrctDbtTxInf) ?", "verbose", LogHandle		
		iTrxCount = ConsoleInput()
		If IsNumeric(iTrxCount) Then
			GetTrxFileCount = Array("",iTrxCount)
		Else
			ConsoleOutput "", "verbose", LogHandle
			ConsoleOutput "<ERROR> INVALID INPUT! PLEASE TRY AGAIN ...", "verbose", LogHandle
			GetTrxFileCount = False
		End If
	Case "3"	'PACS008
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "NOTE: TOTAL TRANSACTIONS  = [No. of 'CdtTrfTxInf'] ", "nolog", LogHandle		
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "SPECIFY NUMBER OF CREDIT TRANSACTIONS (CdtTrfTxInf) ?", "verbose", LogHandle		
		iTrxCount = ConsoleInput()
		If IsNumeric(iTrxCount) Then
			GetTrxFileCount = Array("",iTrxCount)
		Else
			ConsoleOutput "", "verbose", LogHandle
			ConsoleOutput "<ERROR> INVALID INPUT! PLEASE TRY AGAIN ...", "verbose", LogHandle
			GetTrxFileCount = False
		End If
	Case "4"	'PAIN001
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "NOTE: TOTAL TRANSACTIONS  = [No. Of 'PmtInf'] x [No. of 'CdtTrfTxInf'] ", "nolog", LogHandle		
		ConsoleOutput "-----------------------------------------------------------------------", "nolog", LogHandle
		ConsoleOutput "", "verbose", LogHandle
		ConsoleOutput "SPECIFY NUMBER OF PAYMENT INSTRUCTION INFO BLOCKS (PmtInf) ?", "verbose", LogHandle
		iPmtInfCount = ConsoleInput()
		ConsoleOutput "SPECIFY NUMBER OF CREDIT TRANSACTIONS (CdtTrfTxInf) ?", "verbose", LogHandle		
		iTrxCount = ConsoleInput()
		If IsNumeric(iPmtInfCount) And IsNumeric(iTrxCount) Then
			GetTrxFileCount = Array(iPmtInfCount,iTrxCount)
		Else
			ConsoleOutput "", "verbose", LogHandle
			ConsoleOutput "<ERROR> INVALID INPUT! PLEASE TRY AGAIN ...", "verbose", LogHandle
			GetTrxFileCount = False
		End If
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

Dim iCount, iInstAmount, iTotalDbtTrx
Dim nNbOfTxs, nCtrlSum

iTotalDbtTrx = iNumPmtBlocks*PmtCounter

' GROUP HEADER NODES UPDATED HERE ...
Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:MsgId")
strHdrNodes.Text = "MsgID" & GetRandomChars()
Set strHdrNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:NbOfTxs")
strHdrNodes.Text = CLng(iTotalDbtTrx)
Set strHdrNodes = Nothing

'GET DDTRXAMOUNT FOR SINGLE PAYMENT
Set strDbtTxInfNodes = GetSingleNode(ObjXMLDoc,True,"//ns:PmtInf/ns:DrctDbtTxInf/ns:InstdAmt")
iInstAmount = CCur(strDbtTxInfNodes.Text)
Set strDbtTxInfNodes = Nothing
	
Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:CtrlSum")
strHdrNodes.Text = CCur(iInstAmount * iTotalDbtTrx)
Set strHdrNodes = Nothing

'PMNTINFO NODE CLONED HERE ...
Set sPmtInfoCpy1 = GetSingleNode(ObjXMLDoc,True,"//ns:PmtInf").CloneNode(True)
Set stemps = ObjXMLDoc.selectNodes("//ns:PmtInf")
stemps.RemoveAll
Set stemps = Nothing

'CHECK AND SET PMNTINFO (OPTIONAL) SUB-NODE PRESENT OR NOT ...
Set strPmtInfNodes = GetSingleNode(sPmtInfoCpy1,False,"//ns:NbOfTxs")
If Not(strPmtInfNodes Is Nothing) Then
	strPmtInfNodes.Text = CLng(PmtCounter)
End If
Set strHdrNodes = Nothing

Set strPmtInfNodes = GetSingleNode(sPmtInfoCpy1,False,"//ns:CtrlSum")
If Not(strPmtInfNodes Is Nothing) Then
	strPmtInfNodes.Text = CLng(iInstAmount * PmtCounter)
End If
Set strPmtInfNodes = Nothing


For iNum = 1 To iNumPmtBlocks
	
	Call ConsoleOutput("=========" & " STARTED GENERATING PAYMENT INSTRUCTION BLOCK [PmtInf] " & "=========", "verbose", LogHandle)
	Set strPmtInfo1 = sPmtInfoCpy1.CloneNode(True)
	
	''CLONED PMNTINFO EDITED HERE ...
	Set strPmtInfNodes = GetSingleNode(strPmtInfo1,True,"//ns:PmtInfId")
	strPmtInfNodes.Text = "PmtID" & GetRandomChars()
	Set strPmtInfNodes = Nothing

	''DDTRX NODE CLONED HERE ...
	Set sDDTrxCpy1 = GetSingleNode(strPmtInfo1,True,"//ns:DrctDbtTxInf").CloneNode(True)
	Set stemps = strPmtInfo1.selectNodes("//ns:DrctDbtTxInf")
	stemps.RemoveAll
	Set stemps = Nothing
	
	For iCount = 1 To PmtCounter
		If iCount = 1 Then
			Call ConsoleOutput(" ========" & " STARTED GENERATING DEBIT TRANSACTIONS [DrctDbtTxInf] " & "========", "verbose", LogHandle)
			Call ConsoleOutput("Started Generating " & PmtCounter & " Debit Transactions at " & Now, "verbose", LogHandle)
		End If

		Set sTempNode = sDDTrxCpy1.CloneNode(True)
		Set sEndToEndID = GetSingleNode(sTempNode,True,"//ns:EndToEndId")
			sEndToEndID.Text = iCount & GetRandomChars()
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
	
	'PMNTINFO DOC FRAG APPENDED TO OBJXML HERE ...
	GetSingleNode(ObjXMLDoc,True,"//ns:CstmrDrctDbtInitn").AppendChild oDocFrag
	Set oDocFrag = Nothing

Next


End Function	

'###########################################################################

Public Function GeneratePACS008 (ByRef ObjXMLDoc, iNumPmtBlocks, PmtCounter)
Dim iCount, iAmount

Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:MsgId")
strHdrNodes.Text = "MsgID" & GetRandomChars()
Set strHdrNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:NbOfTxs")
strHdrNodes.Text = CLng(PmtCounter)	'CInt(strHdrNodes.Text) + CLng(PmtCounter)
Set strHdrNodes = Nothing

Set strCdtrTxInfNodes = GetSingleNode(ObjXMLDoc,True,"//ns:CdtTrfTxInf/ns:IntrBkSttlmAmt")
iAmount = CCur(strCdtrTxInfNodes.Text)
Set strCdtrTxInfNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,False,"//ns:CtrlSum")
If Not(strHdrNodes Is Nothing) Then
	strHdrNodes.Text = CCur(iAmount * PmtCounter)
End If
Set strHdrNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,False,"//ns:TtIntrBkSttlmAmt")
If Not(strHdrNodes Is Nothing) Then
	strHdrNodes.Text = CCur(iAmount * PmtCounter)
End If
Set strHdrNodes = Nothing

Set sSelectedNode = GetSingleNode(ObjXMLDoc,True,"//ns:CdtTrfTxInf").CloneNode(True)

Set stemps = ObjXMLDoc.selectNodes("//ns:CdtTrfTxInf")
stemps.RemoveAll
Set stemps = Nothing

Set oDocFrag = ObjXMLDoc.CreateDocumentFragment

For iCount = 1 To PmtCounter

	If iCount = 1 Then
		ConsoleOutput "", "verbose", LogHandle
		Call ConsoleOutput(" ========" & " STARTED GENERATING CREDIT TRANSACTIONS [CdtTrfTxInf] " & "========", "verbose", LogHandle)
		Call ConsoleOutput("Started Generating " & PmtCounter & " Credit Transactions at " & Now, "verbose", LogHandle)
	End If
	
	strE2EIds = Left(GetRandomChars(),4) & "RNDID" & iCount
	
	Set sTempNode = sSelectedNode.CloneNode(True)

	Set sEndToEndID = GetSingleNode(sTempNode,False,"//ns:PmtId/ns:InstrId")
	If Not(sEndToEndID Is Nothing) Then
		sEndToEndID.Text = strE2EIds
	End If
	Set sEndToEndID = Nothing
	
	Set sEndToEndID = GetSingleNode(sTempNode,True,"//ns:PmtId/ns:EndToEndId")
		sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing

	Set sEndToEndID = GetSingleNode(sTempNode,True,"//ns:PmtId/ns:TxId")
		sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing

	oDocFrag.AppendChild sTempNode
	Set sTempNode = Nothing

	Call ConsoleOutput ("Generating Payment ... " & iCount, "nolog", LogHandle)
	
	if iCount = CLng(PmtCounter) Then
		Call ConsoleOutput ("All " & PmtCounter & " Payments Generated Successfully at " & Now, "verbose", LogHandle)
		Call ConsoleOutput ("=============" & " COMPLETED TRANSACTION GENERATION [CdtTrfTxInf]" & "============", "verbose", LogHandle)
		ConsoleOutput "", "verbose", LogHandle
	End If
Next


Set sSelectedNode = Nothing
GetSingleNode(ObjXMLDoc,True,"//ns:FIToFICstmrCdtTrf").AppendChild oDocFrag
Set oDocFrag = Nothing

End Function


'###########################################################################

Public Function GeneratePACS003 (ByRef ObjXMLDoc, iNumPmtBlocks, PmtCounter)

Dim iCount, iAmount
Dim strHdrNodes

Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:MsgId")
strHdrNodes.Text = "MsgID" & GetRandomChars()
Set strHdrNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:NbOfTxs")
strHdrNodes.Text = CLng(PmtCounter)	'CInt(strHdrNodes.Text) + CLng(PmtCounter)
Set strHdrNodes = Nothing

Set strDbtTxInfNodes = GetSingleNode(ObjXMLDoc,True,"//ns:DrctDbtTxInf/ns:IntrBkSttlmAmt")
iAmount = CCur(strDbtTxInfNodes.Text)
Set strDbtTxInfNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,False,"//ns:CtrlSum")
If Not(strHdrNodes Is Nothing) Then
	strHdrNodes.Text = CCur(iAmount * PmtCounter)
End If
Set strHdrNodes = Nothing

Set strHdrNodes = GetSingleNode(ObjXMLDoc,False,"//ns:TtIntrBkSttlmAmt")
If Not(strHdrNodes Is Nothing) Then
	strHdrNodes.Text = CCur(iAmount * PmtCounter)
End If
Set strHdrNodes = Nothing

Set sSelectedNode = GetSingleNode(ObjXMLDoc,True,"//ns:DrctDbtTxInf").CloneNode(True)

Set stemps = ObjXMLDoc.selectNodes("//ns:DrctDbtTxInf")
stemps.RemoveAll
Set stemps = Nothing

Set oDocFrag = ObjXMLDoc.CreateDocumentFragment

For iCount = 1 To PmtCounter

	If iCount = 1 Then
		ConsoleOutput "", "verbose", LogHandle
		Call ConsoleOutput(" ========" & " STARTED GENERATING DEBIT TRANSACTIONS [DrctDbtTxInf] " & "========", "verbose", LogHandle)
		Call ConsoleOutput("Started Generating " & PmtCounter & " Debit Transactions at " & Now, "verbose", LogHandle)
	End If
	
	strE2EIds = Left(GetRandomChars(),4) & "RNDID" & iCount
	
	Set sTempNode = sSelectedNode.CloneNode(True)
	
	Set sEndToEndID = GetSingleNode(sTempNode,False,"//ns:PmtId/ns:InstrId")
	If Not(sEndToEndID Is Nothing) Then
		sEndToEndID.Text = strE2EIds
	End If
	Set sEndToEndID = Nothing
	
	Set sEndToEndID = GetSingleNode(sTempNode,True,"//ns:PmtId/ns:EndToEndId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing

	Set sEndToEndID = GetSingleNode(sTempNode,True,"//ns:PmtId/ns:TxId")
	sEndToEndID.Text = strE2EIds
	Set sEndToEndID = Nothing
	
	oDocFrag.AppendChild sTempNode
	Set sTempNode = Nothing

	Call ConsoleOutput ("Generating Debit Transaction ... " & iCount, "nolog", LogHandle)

	if iCount = CLng(PmtCounter) Then
		Call ConsoleOutput ("All " & PmtCounter & " Payments Generated Successfully at " & Now, "verbose", LogHandle)
		Call ConsoleOutput ("=============" & " COMPLETED TRANSACTION GENERATION [DrctDbtTxInf]" & "============", "verbose", LogHandle)
		ConsoleOutput "", "verbose", LogHandle
	End If

Next

Set sSelectedNode = Nothing
GetSingleNode(ObjXMLDoc,True,"//ns:FIToFICstmrDrctDbt").AppendChild oDocFrag
Set oDocFrag = Nothing

End Function

'###########################################################################

Public Function GeneratePAIN001 (ByRef ObjXMLDoc, iNumPmtBlocks, PmtCounter)

	Dim iCount, iInstAmount, iTotalDbtTrx
	Dim nNbOfTxs, nCtrlSum
	
	iTotalDbtTrx = iNumPmtBlocks*PmtCounter

	' GROUP HEADER NODES UPDATED HERE ...
	Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:MsgId")
	strHdrNodes.Text = "MsgID" & GetRandomChars()
	Set strHdrNodes = Nothing
	
	Set strHdrNodes = GetSingleNode(ObjXMLDoc,True,"//ns:NbOfTxs")
	strHdrNodes.Text = CLng(iTotalDbtTrx)
	Set strHdrNodes = Nothing

	'GET CRDTTRXAMOUNT FOR SINGLE PAYMENT
	Set strCrdtTxInfNodes = GetSingleNode(ObjXMLDoc,True,"//ns:PmtInf/ns:CdtTrfTxInf/ns:Amt")
	Set ObjOptionalNode = GetChoiceNode(strCrdtTxInfNodes,arrPAIN001AmtOption)
	iInstAmount = ObjOptionalNode.Text
	Set strCrdtTxInfNodes = Nothing
	Set ObjOptionalNode = Nothing

	Set strHdrNodes = GetSingleNode(ObjXMLDoc,False,"//ns:CtrlSum")
	If Not(strHdrNodes Is Nothing) Then
		strHdrNodes.Text = CCur(iInstAmount * iTotalDbtTrx)
	End If
	Set strHdrNodes = Nothing
		
	'PMNTINFO NODE CLONED HERE ...
	Set sPmtInfoCpy1 = GetSingleNode(ObjXMLDoc,True,"//ns:PmtInf").CloneNode(True)
	Set stemps = ObjXMLDoc.selectNodes("//ns:PmtInf")
	stemps.RemoveAll
	Set stemps = Nothing
	
	'CHECK AND SET PMNTINFO (OPTIONAL) SUB-NODE PRESENT OR NOT ...
	Set strPmtInfNodes = GetSingleNode(sPmtInfoCpy1,False,"//ns:NbOfTxs")
	If Not(strPmtInfNodes Is Nothing) Then
		strPmtInfNodes.Text = CLng(PmtCounter)
	End If
	Set strHdrNodes = Nothing
	
	Set strPmtInfNodes = GetSingleNode(sPmtInfoCpy1,False,"//ns:CtrlSum")
	If Not(strPmtInfNodes Is Nothing) Then
		strPmtInfNodes.Text = CLng(iInstAmount * PmtCounter)
	End If
	Set strPmtInfNodes = Nothing
	
	
	For iNum = 1 To iNumPmtBlocks
		
		Call ConsoleOutput("=========" & " STARTED GENERATING PAYMENT INSTRUCTION BLOCK [PmtInf] " & "=========", "verbose", LogHandle)
		Set strPmtInfo1 = sPmtInfoCpy1.CloneNode(True)
		
		''CLONED PMNTINFO EDITED HERE ...
		Set strPmtInfNodes = GetSingleNode(strPmtInfo1,True,"//ns:PmtInfId")
		strPmtInfNodes.Text = "PmtID" & GetRandomChars()
		Set strPmtInfNodes = Nothing
	
		''CREDITTRX NODE CLONED HERE ...
		Set sDDTrxCpy1 = GetSingleNode(strPmtInfo1,True,"//ns:CdtTrfTxInf").CloneNode(True)
		Set stemps = strPmtInfo1.selectNodes("//ns:CdtTrfTxInf")
		stemps.RemoveAll
		Set stemps = Nothing
		
		For iCount = 1 To PmtCounter
			If iCount = 1 Then
				Call ConsoleOutput(" ========" & " STARTED GENERATING CREDIT TRANSACTIONS [CdtTrfTxInf] " & "========", "verbose", LogHandle)
				Call ConsoleOutput("Started Generating " & PmtCounter & " Credit Transactions at " & Now, "verbose", LogHandle)
			End If
	
			Set sTempNode = sDDTrxCpy1.CloneNode(True)
			Set sEndToEndID = GetSingleNode(sTempNode,True,"//ns:EndToEndId")
				sEndToEndID.Text = iCount & GetRandomChars()
			Set sEndToEndID = Nothing
			
			strPmtInfo1.AppendChild sTempNode
			Set sTempNode = Nothing
			
			Call ConsoleOutput ("Generating Credit Transaction ... " & iCount, "nolog", LogHandle)
			
			If iCount = CLng(PmtCounter) Then
				Call ConsoleOutput ("All " & PmtCounter & " Payments Generated Successfully at " & Now, "verbose", LogHandle)
				Call ConsoleOutput ("=============" & " COMPLETED TRANSACTION GENERATION [CdtTrfTxInf]" & "============", "verbose", LogHandle)
			End If
		Next
		
		'PMNTINFO NODE APPENDED TO THE DOC FRAG HERE ...
		Set oDocFrag = ObjXMLDoc.CreateDocumentFragment
		oDocFrag.AppendChild strPmtInfo1
		
		Call ConsoleOutput("========" & " COMPLETED GENERATING PAYMENT INSTRUCTION BLOCK [PmtInf] " & "========" & vbCrLf & vbCrLf & vbCrLf, "verbose", LogHandle)
	
		Set strPmtInfo1 = Nothing
		Set sDDTrxCpy1 = Nothing
		
		'PMNTINFO DOC FRAG APPENDED TO OBJXML HERE ...
		GetSingleNode(ObjXMLDoc,True,"//ns:CstmrCdtTrfInitn").AppendChild oDocFrag
		Set oDocFrag = Nothing
	
	Next
	
	
End Function	
	

'###########################################################################


Public Function GetSingleNode (ObjXMLDomNode, IsThrowErr, strXPathString)

Dim ObjTempNode 
Set ObjTempNode = ObjXMLDomNode.selectSingleNode(strXPathString)

If Not(ObjTempNode is Nothing) Then
	Set GetSingleNode = ObjTempNode
Else
	Select Case IsThrowErr
		Case False
			Set GetSingleNode = Nothing		
		Case True
			Call ConsoleOutput ("<ERROR> INVALID XML. NODE NOT FOUND ! : " & strXPathString, "verbose", LogHandle)
			If IsReloadExit("") Then
				Call StartSEPABulkGen()
			Else
				ExitApp()
			End If
	End Select
End If

End Function

'###########################################################################

Public Function GetChoiceNode (ObjParentDomNode, arrXPathString)

Dim ObjTempNode, iCount

iCount = 0
For each strXPath in arrXPathString
	msgbox strXPath
	Set ObjTempNode = GetSingleNode (ObjParentDomNode, False, strXPath)
	If Not(ObjTempNode Is Nothing) Then
		Set GetChoiceNode = ObjTempNode
		iCount = iCount + 1
		msgbox "hi"
	End if
Next

If (iCount = 0) Then
	GetChoiceNode = Nothing
End if

End Function

'###########################################################################

Public Function GetRandomChars ()

Set ObjFSO = CreateObject("Scripting.FileSystemObject")

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
WScript.StdOut.WriteLine VBTab & vbTab & "   " & "sepa-Iso2K22-BulkGen version 1.0.5"
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
WScript.StdOut.WriteLine "2. DD CLEARING SETTLEMENT PACS.003"
WScript.StdOut.WriteLine "3. CREDIT CLEARING SETTLEMENT PACS.008"
WScript.StdOut.WriteLine "4. CREDIT TRANSFER INITIATION PAIN.001"
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

End Function

'###########################################################################

Public Function ShowXML()

'Display XML on the console using the XMLDoc.xml property

End Function

'###########################################################################

Public Function SetValue(ObjXML, XPathQuery, strValue)

End Function

'###########################################################################


























