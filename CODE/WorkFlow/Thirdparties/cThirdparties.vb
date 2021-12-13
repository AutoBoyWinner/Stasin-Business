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
	Const FIELD_STRING As String = "name|nameAlias|country|address|phone|email|codeClient|codeFournisseur"
	Const ITEM_STRING As String = "name|name_alias|country|address|phone|email|code_client|code_fournisseur"
	
    Sub Form_OnLoad(ByVal eventData As MFPEventData)
		'TODO add code here to execute when the form is first shown
		Call GetThirpartyList(eventData)
    End Sub

    Sub Form_OnSubmit(ByVal eventData As MFPEventData)
        'TODO add code here to execute when the user presses OK in the form
    End Sub

	Sub thirdpartyList_OnChange(ByVal eventData As MFPEventData) 'TODO change <fieldName> to desired field name
		'TODO add code here to execute when field value of <fieldName> is changed
		Dim thirdpartyList As ListField = eventData.Form.Fields.GetField("thirdpartyList")
		Dim id As String = thirdpartyList.GetSelectedItem().Value
		
		Call SetFieldValue(eventData, id)	
	End Sub
	
	Function SetFieldValue(ByVal eventData As MFPEventData, ByVal id As String)
		
		
		Dim url As String = "https://stasinpierre.with2.dolicloud.com/api/index.php/thirdparties/" & id
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "GET"
		request.ContentType = "application/json"
		request.Headers.Add("DOLAPIKEY", SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		json = "[" & json & "]"
		
		If Not response Is Nothing Then response.Close()
	
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
				
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		'Const FIELD_STRING As String = "name|nameAlias|country|address|phone|email|codeClient|codeFournisseur"
		'Const ITEM_STRING As String = "name|name_alias|country|address|phone|email|code_client|code_fournisseur"
		Dim FieldNameArray() As String = FIELD_STRING.Split("|")
		Dim ItemNameArray() As String = ITEM_STRING.Split("|")
		Dim field As TextField
		For index As Integer = 0 To results.Count-1
			results1 = results.Item(index)			
			
			For j As Integer = 0 To ItemNameArray.Length - 1
				field = eventData.Form.Fields.GetField(FieldNameArray(j))
				field.Value = results1.Item(ItemNameArray(j))
			Next
		Next
	End Function
	
	Function GetThirpartyList(ByVal eventData As MFPEventData)
		Dim url As String = "https://stasinpierre.with2.dolicloud.com/api/index.php/thirdparties?sortfield=t.rowid&sortorder=ASC"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "GET"
		request.ContentType = "application/json"
		request.Headers.Add("DOLAPIKEY", SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		
		If Not response Is Nothing Then response.Close()
	
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		Dim documentList As ListField = eventData.Form.Fields.GetField("thirdpartyList")

		documentList.FindMode = False
		documentList.Items.Clear()
		
		
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		Dim listItem As ListItem 
		documentList.MaxSearchResults = results.Count + 1

		For index As Integer = 0 To results.Count - 1
			results1 = results.Item(index)			
			listItem = New ListItem(results1.Item("name") & "| " & results1.Item("code_client"), results1.Item("id"))
			'listItem = New ListItem(results.Count.ToString & "| " & results1.Item("code_client"), results1.Item("id").ToString)
			documentList.Items.Add(listItem)
			'Next
		Next		
		
		GetThirpartyList = response.StatusCode
	End Function

End Module
