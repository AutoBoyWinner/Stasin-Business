Const SECURITY_TOKEN = "fe6bc0f5c890e445397a90c2cf1ca319c40d1857"
Sub Form_OnLoad(Form)
	Form.SetFieldValue "status2", "test.."
	Form.StatusMsg("Start")	
	
	Form.SetFieldValue "modulepart", "invoice"
	
	Form.SetFieldValue "ref", "FA2101-0002"
	Form.SetFieldValue "downloadPath", "C:\"
End Sub 


Function Form_OnValidate(Form)

End Function


Sub Field_OnChange(Form, FieldName, FieldValue)

End Sub


Function Field_OnValidate(FieldName, FieldValue)

End Function


Sub Button_OnClick(Form, ButtonName)
	
	If (ButtonName = "searchDoc") Then
		Dim modulepart: modulepart = Form.GetFieldValue("modulepart")
		Dim ref: ref = Form.GetFieldValue("ref")
		If(modulepart <> "") And (ref <>"") Then
			Call GetDocList(Form, modulepart, ref)
			'Form.SetFieldValue "status2", resp
		Else
			Form.SetFieldValue "status2", "modulepart or ref is empty!"			
		End If
	End If
	If (ButtonName = "download") Then
		Dim downloadPath: downloadPath = Form.GetFieldValue("downloadPath") 
		Call DownloadDoc(Form, downloadPath)
		If downloadPath <> "" Then
			
		Else
			Form.SetFieldValue "status2", "downloadPath is empty!"		
		End If			
	End If
	If (ButtonName = "delete") Then
		Call DeleteDoc(Form)	
	End If
End Sub
Function DeleteDoc(Form)
	Form.SetFieldValue "status2", "deleting......."
	
	Dim selectedItemValue: selectedItemValue = Form.GetFieldValue("docList") 
	
	Dim valueArray: valueArray  = Split(selectedItemValue,"|")
	Dim modulepart, ref, srcFileName, fullDownloadPath, original_file
	
	modulepart = valueArray(0)
	ref =valueArray(1)
	srcFileName = valueArray(2)
	
	fullDownloadPath = downloadPath & srcFileName
	'original_file = ref & "/" & srcFileName
	
	
	Dim HTTP, url

	Set HTTP = CreateObject("Microsoft.XMLHTTP")

	url = "https://stasinpierre.with2.dolicloud.com/api/index.php/documents?modulepart=" & modulepart & "&original_file=" & ref & "%2F" & srcFileName 
	'Form.SetFieldValue "status2", url

	HTTP.Open "DELETE", url, False
	HTTP.setRequestHeader "Content-Type", "application/json"
	HTTP.setRequestHeader "DOLAPIKEY", SECURITY_TOKEN
	HTTP.send

	Dim response_status
	
	response_status =  HTTP.status
	Dim Mymessage
	If response_status = "200" Then
		Mymessage = "success deleteing file...." & srcFileName
	ElseIf response_status = "404" Then
		Mymessage = "file not found..."
	Else
		Mymessage = "error occurs..."
	End If
	Form.SetFieldValue "status2", Mymessage
End Function
Function DownloadDoc(Form, downloadPath)
	Form.SetFieldValue "status2", "downloading......."
	Dim selectedItemValue: selectedItemValue = Form.GetFieldValue("docList") 
	
	Dim valueArray: valueArray  = Split(selectedItemValue,"|")
	Dim modulepart, ref, srcFileName, fullDownloadPath, original_file
	
	modulepart = valueArray(0)
	ref =valueArray(1)
	srcFileName = valueArray(2)
	
	fullDownloadPath = downloadPath & srcFileName
	'original_file = ref & "/" & srcFileName
	
	
	Dim HTTP, url

	Set HTTP = CreateObject("Microsoft.XMLHTTP")

	url = "https://stasinpierre.with2.dolicloud.com/api/index.php/documents/download?modulepart=" & modulepart & "&original_file=" & ref & "%2F" & srcFileName 
	

	HTTP.Open "GET", url, False
	HTTP.setRequestHeader "Content-Type", "application/json"
	HTTP.setRequestHeader "DOLAPIKEY", SECURITY_TOKEN
	HTTP.send
	
	Dim response_status: response_status =  HTTP.status
	If response_status = "200" Then
		Dim responseText
	
		responseText =  HTTP.responseText
		
		Set json = New VbsJson		
		
		Set results = json.Decode(responseText)
		Dim filecontent : filecontent = results("content")
		
		Dim outByteArray: outByteArray = decodeBase64(filecontent)
		Call	WriteBinaryFile(fullDownloadPath, outByteArray)
		Form.SetFieldValue "status2", "success download..." & fullDownloadPath
	ElseIf response_status = "404" Then
		Form.SetFieldValue "status2", "file not found..."
	Else
		Form.SetFieldValue "status2", "error occurs"
	End If
End Function
Function GetDocList(Form, modulepart, ref)
	Dim HTTP, url

	Set HTTP = CreateObject("Microsoft.XMLHTTP")

	url = "https://stasinpierre.with2.dolicloud.com/api/index.php/documents?modulepart=" & modulepart & "&ref=" & ref
	

	HTTP.Open "GET", url, False
	HTTP.setRequestHeader "Content-Type", "application/json"
	HTTP.setRequestHeader "DOLAPIKEY", SECURITY_TOKEN
	HTTP.send

	Dim response_status: response_status =  HTTP.status
	If response_status = "200" Then
		Dim responseText
	
		responseText =  "{""next"":""ex"",""patients"":" & HTTP.responseText & "}"
		
		Set json = New VbsJson
		
		
		Set results = json.Decode(responseText)
		Dim patients : patients = results("patients")
		Call Form.Fields.Field("docList").AddListItem("", "")
		
		For Each patient In patients
			Call Form.Fields.Field("docList").AddListItem(patient("name"), modulepart & "|" & ref & "|" & patient("name"))
		Next
	ElseIf response_status = "404" Then
		Form.SetFieldValue "status2", "file not found..."
	Else
		Form.SetFieldValue "status2", "error occurs"
	End If
	
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

Class VbsJson
	'Author: Demon
	'Date: 2012/5/3
	'Website: http://demon.tw
	Private Whitespace, NumberRegex, StringChunk
	Private b, f, r, n, t

	Private Sub Class_Initialize
		Whitespace = " " & vbTab & vbCr & vbLf
		b = ChrW(8)
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab

		Set NumberRegex = New RegExp
		NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
		NumberRegex.Global = False
		NumberRegex.MultiLine = True
		NumberRegex.IgnoreCase = True

		Set StringChunk = New RegExp
		StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
		StringChunk.Global = False
		StringChunk.MultiLine = True
		StringChunk.IgnoreCase = True
	End Sub

	'Return a JSON string representation of a VBScript data structure
	'Supports the following objects and types
	'+-------------------+---------------+
	'| VBScript          | JSON          |
	'+===================+===============+
	'| Dictionary        | object        |
	'+-------------------+---------------+
	'| Array             | array         |
	'+-------------------+---------------+
	'| String            | string        |
	'+-------------------+---------------+
	'| Number            | number        |
	'+-------------------+---------------+
	'| True              | true          |
	'+-------------------+---------------+
	'| False             | false         |
	'+-------------------+---------------+
	'| Null              | null          |
	'+-------------------+---------------+
	Public Function Encode(ByRef obj)
		Dim buf, i, c, g
		Set buf = CreateObject("Scripting.Dictionary")
		Select Case VarType(obj)
			Case vbNull
				buf.Add buf.Count, "null"
			Case vbBoolean
				If obj Then
					buf.Add buf.Count, "true"
				Else
					buf.Add buf.Count, "false"
				End If
			Case vbInteger, vbLong, vbSingle, vbDouble
				buf.Add buf.Count, obj
			Case vbString
				buf.Add buf.Count, """"
				For i = 1 To Len(obj)
					c = Mid(obj, i, 1)
					Select Case c
						Case """" buf.Add buf.Count, "\"""
						Case "\"  buf.Add buf.Count, "\\"
						Case "/"  buf.Add buf.Count, "/"
						Case b    buf.Add buf.Count, "\b"
						Case f    buf.Add buf.Count, "\f"
						Case r    buf.Add buf.Count, "\r"
						Case n    buf.Add buf.Count, "\n"
						Case t    buf.Add buf.Count, "\t"
						Case Else
							If AscW(c) >= 0 And AscW(c) <= 31 Then
								c = Right("0" & Hex(AscW(c)), 2)
								buf.Add buf.Count, "\u00" & c
							Else
								buf.Add buf.Count, c
							End If
					End Select
				Next
				buf.Add buf.Count, """"
			Case vbArray + vbVariant
				g = True
				buf.Add buf.Count, "["
				For Each i In obj
					If g Then g = False Else buf.Add buf.Count, ","
					buf.Add buf.Count, Encode(i)
				Next
				buf.Add buf.Count, "]"
			Case vbObject
				If TypeName(obj) = "Dictionary" Then
					g = True
					buf.Add buf.Count, "{"
					For Each i In obj
						If g Then g = False Else buf.Add buf.Count, ","
						buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
					Next
					buf.Add buf.Count, "}"
				Else
					Err.Raise 8732,,"None dictionary object"
				End If
			Case Else
				buf.Add buf.Count, """" & CStr(obj) & """"
		End Select
		Encode = Join(buf.Items, "")
	End Function

	'Return the VBScript representation of ``str(``
	'Performs the following translations in decoding
	'+---------------+-------------------+
	'| JSON          | VBScript          |
	'+===============+===================+
	'| object        | Dictionary        |
	'+---------------+-------------------+
	'| array         | Array             |
	'+---------------+-------------------+
	'| string        | String            |
	'+---------------+-------------------+
	'| number        | Double            |
	'+---------------+-------------------+
	'| true          | True              |
	'+---------------+-------------------+
	'| false         | False             |
	'+---------------+-------------------+
	'| null          | Null              |
	'+---------------+-------------------+
	Public Function Decode(ByRef str)
		Dim idx
		idx = SkipWhitespace(str, 1)

		If Mid(str, idx, 1) = "{" Then
			Set Decode = ScanOnce(str, 1)
		Else
			Decode = ScanOnce(str, 1)
		End If
	End Function

	Private Function ScanOnce(ByRef str, ByRef idx)
		Dim c, ms

		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)

		If c = "{" Then
			idx = idx + 1
			Set ScanOnce = ParseObject(str, idx)
			Exit Function
		ElseIf c = "[" Then
			idx = idx + 1
			ScanOnce = ParseArray(str, idx)
			Exit Function
		ElseIf c = """" Then
			idx = idx + 1
			ScanOnce = ParseString(str, idx)
			Exit Function
		ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = Null
			Exit Function
		ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ScanOnce = True
			Exit Function
		ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
			idx = idx + 5
			ScanOnce = False
			Exit Function
		End If

		Set ms = NumberRegex.Execute(Mid(str, idx))
		If ms.Count = 1 Then
			idx = idx + ms(0).Length
			ScanOnce = CDbl(ms(0))
			Exit Function
		End If

		Err.Raise 8732,,"No JSON object could be ScanOnced"
	End Function

	Private Function ParseObject(ByRef str, ByRef idx)
		Dim c, key, value
		Set ParseObject = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)

		If c = "}" Then
			Exit Function
		ElseIf c <> """" Then
			Err.Raise 8732,,"Expecting property name"
		End If

		idx = idx + 1

		Do
			key = ParseString(str, idx)

			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) <> ":" Then
				Err.Raise 8732,,"Expecting : delimiter"
			End If


			idx = SkipWhitespace(str, idx + 1)

			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If


			' Added if for possible empty array
			If Mid(str, idx, 1) = "]" Then
				value = ""
				idx = idx + 1
			End If

			ParseObject.Add key, value

			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "}" Then
				Exit Do
			ElseIf c <> "," Then
				Err.Raise 8732,,"Expecting , delimiter"
			End If

			idx = SkipWhitespace(str, idx + 1)
			c = Mid(str, idx, 1)
			If c <> """" Then
				Err.Raise 8732,,"Expecting property name"
			End If

			idx = idx + 1
		Loop

		idx = idx + 1
	End Function

	Private Function ParseArray(ByRef str, ByRef idx)
		Dim c, values, value
		Set values = CreateObject("Scripting.Dictionary")
		idx = SkipWhitespace(str, idx)
		c = Mid(str, idx, 1)

		If c = "]" Then
			ParseArray = values.Items
			Exit Function
		End If

		Do
			idx = SkipWhitespace(str, idx)
			If Mid(str, idx, 1) = "{" Then
				Set value = ScanOnce(str, idx)
			Else
				value = ScanOnce(str, idx)
			End If
			values.Add values.Count, value

			idx = SkipWhitespace(str, idx)
			c = Mid(str, idx, 1)
			If c = "]" Then
				Exit Do
			ElseIf c <> "," Then
				Err.Raise 8732,,"Expecting , delimiter"
			End If

			idx = idx + 1
		Loop

		idx = idx + 1
		ParseArray = values.Items
	End Function

	Private Function ParseString(ByRef str, ByRef idx)
		Dim chunks, content, terminator, ms, esc, char
		Set chunks = CreateObject("Scripting.Dictionary")

		Do
			Set ms = StringChunk.Execute(Mid(str, idx))
			If ms.Count = 0 Then
				Err.Raise 8732,,"Unterminated string starting"
			End If

			content = ms(0).Submatches(0)
			terminator = ms(0).Submatches(1)
			If Len(content) > 0 Then
				chunks.Add chunks.Count, content
			End If

			idx = idx + ms(0).Length

			If terminator = """" Then
				Exit Do
			ElseIf terminator <> "\" Then
				Err.Raise 8732,,"Invalid control character"
			End If

			esc = Mid(str, idx, 1)

			If esc <> "u" Then
				Select Case esc
					Case """" char = """"
					Case "\"  char = "\"
					Case "/"  char = "/"
					Case "b"  char = b
					Case "f"  char = f
					Case "n"  char = n
					Case "r"  char = r
					Case "t"  char = t
					Case Else Err.Raise 8732,,"Invalid escape"
				End Select
				idx = idx + 1
			Else
				char = ChrW("&H" & Mid(str, idx + 1, 4))
				idx = idx + 5
			End If

			chunks.Add chunks.Count, char
		Loop

		ParseString = Join(chunks.Items, "")
	End Function

	Private Function SkipWhitespace(ByRef str, ByVal idx)
		Do While idx <= Len(str) And _
			InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		SkipWhitespace = idx
	End Function
End Class


