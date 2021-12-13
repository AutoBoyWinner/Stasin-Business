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

Module Script
	Const SECURITY_TOKEN As String = "fe6bc0f5c890e445397a90c2cf1ca319c40d1857" 
	Const MODULEPART_STRING As String = "expensereport|invoice|propal|project|order_supplier"
    Sub Form_OnLoad(ByVal eventData As MFPEventData)
		'TODO add code here to execute when the form is first shown
		Dim status As TextField = eventData.Form.Fields.GetField("status")
		status.Value = "test0"
		Call GetModulepartyList(eventData)
	End Sub

    Sub Form_OnSubmit(ByVal eventData As MFPEventData)
        'TODO add code here to execute when the user presses OK in the form
    End Sub

	Sub modulepart_OnChange(ByVal eventData As MFPEventData) 'TODO change <fieldName> to desired field name
		'TODO add code here to execute when field value of <fieldName> is changed
		'Dim modulepartList As ListField = eventData.Form.Fields.GetField("modulepart")
		
	End Sub
	
	Sub search_OnChange(ByVal eventData As MFPEventData)
		'TODO add code here to execute when field value of <fieldName> is changed
		Dim status As TextField = eventData.Form.Fields.GetField("status")
		status.Value = "test-1"
		
		Dim modulepartList As ListField = eventData.Form.Fields.GetField("modulepart")
		Dim modulepart As String = modulepartList.GetSelectedItem().Value.ToString
		Dim refField As TextField = eventData.Form.Fields.GetField("ref")
		Dim ref As String = refField.Value
		status.Value = "test1"
		Call GetDocumentList(eventData, modulepart, ref)
		status.Value = "test2"
	End Sub

	Function GetModulepartyList(ByVal eventData As MFPEventData)
		Dim modulepartList As ListField = eventData.Form.Fields.GetField("modulepart")
		Dim modulepartArray() As String = MODULEPART_STRING.Split("|")
		
		For index As Integer = 0 To modulepartArray.Length - 1
			Dim listItem As ListItem = New ListItem(modulepartArray(index), modulepartArray(index))
			modulepartList.Items.Add(listItem)
		Next		
	End Function
	Function GetDocumentList(ByVal eventData As MFPEventData, ByVal modulepart As String, ByVal ref As String)
		
		Dim url As String = "https://stasinpierre.with2.dolicloud.com/api/index.php/documents?modulepart=" & modulepart & "&ref=" & ref
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
		Dim responseCode As String = response.StatusCode
	
		'msgbox response.ToString 
	
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		
		If Not response Is Nothing Then response.Close()
	
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		
		
		Dim documentList As ListField = eventData.Form.Fields.GetField("documentList")

		documentList.FindMode = False
		documentList.Items.Clear()

		For index As Integer = 0 To results.Count-1
			Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) = results.Item(index)			
			
			Dim listItem As ListItem = New ListItem(results1.Item("name"), results1.Item("name"))
			documentList.Items.Add(listItem)
			'Next
		Next		
	End Function
	
	
End Module
