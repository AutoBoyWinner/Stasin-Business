'LoadAssembly:System.Web.Extensions.dll

' Be sure to copy the dll System.WebExtensions.dll into the AutoStore directory in "Program Files"
' SEE README.md for more details

Option Strict Off

Imports System
Imports NSi.AutoStore.Capture.DataModel
Imports System.Net
Imports System.Web
Imports System.IO
Imports System.Text
Imports System.Web.Script.Serialization
'Imports Newtonsoft.Json

Module Script

	
	Const SECURITY_TOKEN As String = "fe6bc0f5c890e445397a90c2cf1ca319c40d1857" 
	
	Sub Form_OnLoad(ByVal eventData As MFPEventData)
		'TODO add code here to execute when the form is first shown
			Dim modulepart As LabelField = eventData.Form.Fields.GetField("modulepart")
		modulepart.Text = "propal"
		modulepart.IsHidden = False
		eventData.Form.Fields.GetField("notrigger").Value = "1"
		
		
		Call MyStatusMessage(eventData, "status : ")
		'Call GetModulepartyList(eventData)
		Dim resp As String
		resp = GetDocumentList(eventData)
		'Call MyStatusMessage(eventData, resp)
	End Sub
	Function MyStatusMessage(ByVal eventData As MFPEventData, ByVal msg_string As String)
		Dim status As LabelField = eventData.Form.Fields.GetField("status")
		status.Text = msg_string
		status.IsHidden = False
	End Function
    Sub Form_OnSubmit(ByVal eventData As MFPEventData)
		'TODO add code here to execute when the user presses OK in the form
		'MyStatusMessage(eventData, "ok_submit")
    End Sub

	Sub register_OnChange(ByVal eventData As MFPEventData)
		'TODO add code here to execute when field value of <fieldName> is changed
		Dim ref As ListField = eventData.Form.Fields.GetField("ref")
		Dim proposal_id As String = ref.GetSelectedItem().Text.Split("|")(0)
		
		Dim notrigger As String= eventData.Form.Fields.GetField("notrigger").Value
		
		Dim resp As String = register_proposal(proposal_id, notrigger)
		If resp = "OK" Then
			Call MyStatusMessage(eventData, "status : successed")
		Else
			Call MyStatusMessage(eventData, "status : failed")
		End If
	End Sub

	Function register_proposal(ByVal proposal_id As String, ByVal notrigger As String)
		Dim url As String = "https://stasinpierre.with2.dolicloud.com/api/index.php/proposals/" &  proposal_id  & "/validate"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "POST"
		request.ContentType = "application/json"
		request.Headers.Add("DOLAPIKEY", SECURITY_TOKEN)
		
		Dim json_data As String = "{""notrigger"": " & notrigger & "}"
		Dim json_bytes() As Byte = System.Text.Encoding.ASCII.GetBytes(json_data)
		request.ContentLength = json_bytes.Length
		
    
		
		Dim stream As IO.Stream = request.GetRequestStream

		stream.Write(json_bytes, 0, json_bytes.Length)


		Dim response As HttpWebResponse = request.GetResponse
		Dim status As String = response.StatusCode.ToString
		
		Dim dataStream As IO.Stream = response.GetResponseStream()
		Dim reader As New IO.StreamReader(dataStream)          ' Open the stream using a StreamReader for easy access.
		Dim responseFromServer As String = reader.ReadToEnd()  ' Read the content.

		' Cleanup the streams and the response.
		reader.Close()
		dataStream.Close()
		response.Close()
		register_proposal = status
	End Function
	Function GetDocumentList(ByVal eventData As MFPEventData)
		
		Dim url As String = "https://stasinpierre.with2.dolicloud.com/api/index.php/proposals?sortfield=t.rowid&sortorder=ASC"
		'		If seachKey.Length > 0 Then
		'			url = url & "&thirdparty_ids=" & term
		'		End If
	
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "GET"
		request.ContentType = "application/json"
		'request.Headers.Add("Authorization", "DOLAPIKEY " & SECURITY_TOKEN)
		request.Headers.Add("DOLAPIKEY", SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		'Dim responseCode As String = response.StatusCode
	
		'msgbox response.ToString 
	
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		
		If Not response Is Nothing Then response.Close()
	
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		
		
		Dim documentList As ListField = eventData.Form.Fields.GetField("ref")

		documentList.FindMode = False
		documentList.MaxSearchResults = 100
		documentList.Items.Clear()

		'Dim status As LabelField = eventData.Form.Fields.GetField("status")
		'status.Text = results.Count.ToString & "GGGAAB"
		
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		Dim listItem As ListItem 
		For index As Integer = 0 To results.Count-1
			results1 = results.Item(index)			
			
			'listItem = New ListItem(results1.Item("id").ToString & "|" & results1.Item("ref").ToString & "|" & results1.Item("ref_client").ToString, results1.Item("ref").ToString)
			listItem = New ListItem(results1.Item("id") & "|" & results1.Item("ref") & "| " & results1.Item("ref_client"),  results1.Item("ref"))
			documentList.Items.Add(listItem)
			'Next
		Next	
		
		
		
		GetDocumentList = response.StatusCode.ToString
	End Function
	
	
End Module
