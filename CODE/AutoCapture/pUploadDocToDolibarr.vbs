Const SECURITY_TOKEN  = "fe6bc0f5c890e445397a90c2cf1ca319c40d1857"
Sub UpoadDoc_OnLoad
	EKOManager.StatusMessage ("----------------------------------------modulepart = " & modulepart)
	EKOManager.StatusMessage ("----------------------------------------ref = " & ref)
	
	Set KDocument = KnowledgeObject.GetFirstDocument
	If Not(KDocument Is Nothing) Then
		Set PTopic = KnowledgeObject.GetPersistenceTopic()
		Set Topic  = KnowledgeContent.GetTopicInterface

		If Not(Topic Is Nothing) Then
			
			
			'get Persistence topic to check tag so we don't process same KO again
			Set MyTagEntry = PTopic.GetEntry( "SplitDocument",0 )
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.GetFile(KDocument.FilePath)
			Dim originalFileName : originalFileName = objFSO.GetBaseName(objFile)
			Dim oFileName: oFileName = objFSO.GetFileName(KDocument.FilePath)
			EKOManager.StatusMessage ("------------------------******-- oFileName = " & oFileName)
			EKOManager.StatusMessage ("------------------------******--originalFileName = " & originalFileName)
			EKOManager.StatusMessage ("New document at = " &  KDocument.FilePath)
			Dim resp : resp = SendDocument(oFileName, KDocument.FilePath, modulepart, ref)
			EKOManager.StatusMessage ("resp (should be 200) = " & resp)
		End If
	End If
End Sub

Sub UpoadDoc_OnUnload

End Sub
Function SendDocument(filename, DocPath, modulepart, ref)

	Set HTTP = CreateObject("Microsoft.XMLHTTP")

	url = "https://stasinpierre.with2.dolicloud.com/api/index.php/documents/upload"
	EKOManager.StatusMessage ("SendDocument = " &  DocPath)
	Dim inByteArray: inByteArray = ReadBinaryFile(DocPath)
	
	Dim base64Encoded: base64Encoded = encodeBase64(inByteArray)
	Dim cleanFile : cleanFile = Replace(Trim(CStr(base64Encoded)), vbLf, "")
	
	jsonData = "{""filename"": """ & filename & """, ""modulepart"": """ & modulepart & """,  ""ref"": """ & ref & """,  ""subdir"": """",  ""filecontent"": """ & cleanFile &""",  ""fileencoding"": ""base64"", ""overwriteifexists"": ""1"" }"
	
	'Set fs = CreateObject("Scripting.FileSystemObject")
	'Set objFile = fs.CreateTextFile("C:\AutoStore\KPFTraining\Outputout.pdf", True)
	'objFile.Write cleanFile
	'objFile.Close
	
	'Dim outByteArray: outByteArray = decodeBase64(cleanFile)
	'Dim opp : opp = "C:\2.pdf"
	'Call	WriteBinaryFile(opp, outByteArray)
	
	EKOManager.StatusMessage ("-----------------------OhK")

	HTTP.Open "POST", url, False
	HTTP.setRequestHeader "Content-Type", "application/json"
	
	HTTP.setRequestHeader "DOLAPIKEY", SECURITY_TOKEN
	On Error Resume Next
	HTTP.send jsonData

	
	'msgbox HTTP.responseText
	SendDocument = HTTP.status ' Expect 201, need to fail if not
End Function

Function ReadBinaryFile(FileName)
	Const adTypeBinary = 1

	'Create Stream object
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")

	'Specify stream type - we want To get binary data.
	BinaryStream.Type = adTypeBinary

	'Open the stream
	BinaryStream.Open

	'Load the file data from disk To stream object
	BinaryStream.LoadFromFile FileName

	'Open the stream And get binary data from the object
	ReadBinaryFile = BinaryStream.Read
End Function

Function WriteBinaryFile(FileName, nByteArray)
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2
  
	'Create Stream object
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")
  
	'Specify stream type - we want To save binary data.
	BinaryStream.Type = adTypeBinary
  
	'Open the stream And write binary data To the object
	BinaryStream.Open
	BinaryStream.Write nByteArray
  
	'Save binary data To disk
	BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Private Function encodeBase64(bytes)
	Dim DM, EL
	Set DM = CreateObject("Microsoft.XMLDOM")
	' Create temporary node with Base64 data type
	Set EL = DM.createElement("tmp")
	EL.DataType = "bin.base64"
	' Set bytes, get encoded String
	EL.NodeTypedValue = bytes
	encodeBase64 = EL.Text
End Function

Private Function decodeBase64(filecontent)
	Dim DM, EL
	Set DM = CreateObject("Microsoft.XMLDOM")
	' Create temporary node with Base64 data type
	Set EL = DM.createElement("tmp")
	EL.DataType = "bin.base64"
	EL.Text = filecontent
	decodeBase64 = EL.NodeTypedValue
End Function
