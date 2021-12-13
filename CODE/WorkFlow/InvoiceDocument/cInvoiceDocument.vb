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
		modulepart.Text = "invoice"
		modulepart.IsHidden = False
	
		
		Dim status As LabelField = eventData.Form.Fields.GetField("status")
		status.Text = "test885"
		status.IsHidden = False
		'Call GetModulepartyList(eventData)
		Dim resp As String
		resp = GetDocumentList(eventData)
		'status.Text = resp.ToString
	End Sub

    Sub Form_OnSubmit(ByVal eventData As MFPEventData)
        'TODO add code here to execute when the user presses OK in the form
    End Sub

	Sub register_OnChange(ByVal eventData As MFPEventData)
		'TODO add code here to execute when field value of <fieldName> is changed
		Dim ref As ListField = eventData.Form.Fields.GetField("ref")
		Dim invoice_id As Long = Convert.ToInt64(ref.GetSelectedItem().Text.Split("|")(0))
		
		Dim close_code As String= eventData.Form.Fields.GetField("closeCode").Value
		Dim close_note As String= eventData.Form.Fields.GetField("closeNote").Value
		
		Dim status As LabelField = eventData.Form.Fields.GetField("status")
		
		
		Dim url As String ="https://stasinpierre.with2.dolicloud.com/api/index.php/invoices/" & 6 & "/settopaid"
		'		If seachKey.Length > 0 Then
		'			url = url & "&thirdparty_ids=" & term
		'		End If
		Dim jsonString As String = "{""close_code"":""" & close_code & """, ""close_note"":""" & close_note & """}"
		
		

		
		Dim webClient As WebClient = New WebClient()
		Dim resByte As Byte()
		Dim resString As String
		Dim reqString() As Byte

		Try
			webClient.Headers("content-type") = "application/json"
			webClient.Headers("DOLAPIKEY") =  SECURITY_TOKEN
			reqString = Encoding.Default.GetBytes(jsonString)
			resByte = webClient.UploadData(url, "post", reqString)
			resString = Encoding.Default.GetString(resByte)
			'Console.WriteLine(resString)
			status.Text = "hhh"
			webClient.Dispose()
			'Return True
			status.Text = "sdfsdf"
		Catch ex As Exception
			'Console.WriteLine(ex.Message)
			'	status.Text = "ggg"
		End Try
		'Return False
		
		
	
	End Sub

	
	Function GetDocumentList(ByVal eventData As MFPEventData)
		
		Dim url As String = "https://stasinpierre.with2.dolicloud.com/api/index.php/invoices?sortfield=t.rowid&sortorder=ASC"
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
			listItem = New ListItem(results1.Item("id") & "|" & results1.Item("ref"), results1.Item("id"))
			documentList.Items.Add(listItem)
			'Next
		Next	
		
		
		
		GetDocumentList = response.StatusCode
	End Function
	
	
End Module
