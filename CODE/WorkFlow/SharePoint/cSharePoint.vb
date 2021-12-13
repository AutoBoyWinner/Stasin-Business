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
Imports Microsoft.VisualBasic.Strings
Imports System.Web.Script.Serialization
'Imports Newtonsoft.Json

Module Script
	Const SECURITY_TOKEN As String = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyIsImtpZCI6Im5PbzNaRHJPRFhFSzFqS1doWHNsSFJfS1hFZyJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvc3Rhc2luZGV2LnNoYXJlcG9pbnQuY29tQDFiNDQ3NDdlLTQyM2ItNDE5Ny1iYTI4LTExNmJmYWE2YjU0ZiIsImlzcyI6IjAwMDAwMDAxLTAwMDAtMDAwMC1jMDAwLTAwMDAwMDAwMDAwMEAxYjQ0NzQ3ZS00MjNiLTQxOTctYmEyOC0xMTZiZmFhNmI1NGYiLCJpYXQiOjE2Mjg5MDgwNTksIm5iZiI6MTYyODkwODA1OSwiZXhwIjoxNjI4OTk0NzU5LCJpZGVudGl0eXByb3ZpZGVyIjoiMDAwMDAwMDEtMDAwMC0wMDAwLWMwMDAtMDAwMDAwMDAwMDAwQDFiNDQ3NDdlLTQyM2ItNDE5Ny1iYTI4LTExNmJmYWE2YjU0ZiIsIm5hbWVpZCI6IjFmYjlhYzczLWZmODgtNGZlZS1iMGQyLTZlMmE4ZGMwZGFiNEAxYjQ0NzQ3ZS00MjNiLTQxOTctYmEyOC0xMTZiZmFhNmI1NGYiLCJvaWQiOiI1NmEyMzYxMC1iY2M0LTQ3MzgtYTdkZS0wYTI2MGQ0ZjU1ODgiLCJzdWIiOiI1NmEyMzYxMC1iY2M0LTQ3MzgtYTdkZS0wYTI2MGQ0ZjU1ODgiLCJ0cnVzdGVkZm9yZGVsZWdhdGlvbiI6ImZhbHNlIn0.gAOyYulvcLkWs_p-GK-PWc7iW7-RZ132JY2SeVLEcu_-aNZY9zP_Zu0-7Tv0mnZkY75UrMhTGNI94t0IotAgDTWZil6aRopdvSIhYUg-5DXlLeNfAQK9AY1nGfqKnUh0LaztYR7Tm56h9pvWTAGCR1O2mn43F0L2ACdXm0kp9DM0y-W_P47vpJ40TC7s6cQ5n2wuxtP8GY593uCvek0x6pZc3glZeCffrsdYygUALQpyRcUfJxVkRHAR4uMKvoD9RvWBsdC1-TlbUINQauUf1MPmYRD7aX9pymIXsvez7-IAs4ABQPBBMQHYnL3YhN8YlONibUmeenz1dH-EV010Dw"
	Const FIELD_STRING As String = "nom|prenom|dateofbirth"
	Const ITEM_STRING As String = "nom|prenom|dateofbirth"
	
    Sub Form_OnLoad(ByVal eventData As MFPEventData)
		'TODO add code here to execute when the form is first shown
		Call SetHiddenIDValue(eventData, "-1")
		Call MyStatusMSG(eventData, "start")
		Call GetPatientList(eventData)
	End Sub
	
	Sub Form_OnSubmit(ByVal eventData As MFPEventData)
        'TODO add code here to execute when the user presses OK in the form
    End Sub

	Sub patientList_OnChange(ByVal eventData As MFPEventData) 'TODO change <fieldName> to desired field name
		'TODO add code here to execute when field value of <fieldName> is changed
		Dim documentList As ListField = eventData.Form.Fields.GetField("patientList")
		Dim id As String = documentList.GetSelectedItem().Value
		Call SetHiddenIDValue(eventData, id)
		Call GetPaitentInfoByID(eventData, id)
		Call MyStatusMSG(eventData, "selected [ID: " & id & "] data.")
	End Sub
	
	Sub newBtn_OnChange(ByVal eventData As MFPEventData) 'TODO change <fieldName> to desired field name
		Call InitialFieldValue(eventData)	
		Call MyStatusMSG(eventData, "possible to create a new data.")
	End Sub
	
	Sub updateBtn_OnChange(ByVal eventData As MFPEventData) 'TODO change <fieldName> to desired field name
		Dim nom, prenom, dateofbirth As String
		nom = eventData.Form.Fields.GetField("nom").Value
		prenom = eventData.Form.Fields.GetField("prenom").Value
		dateofbirth = eventData.Form.Fields.GetField("dateofbirth").Value
		If nom <> "" And prenom <> "" And dateofbirth <> "" Then
			If GetHiddenIDValue(eventData) = "-1" Then	'create new data
				If IsExistPatientInfo(eventData, nom & " " & prenom) = False Then
					Call CreatePatientInfo(eventData)
					MyStatusMSG(eventData, "successed creating a data.")
				Else
					MyStatusMSG(eventData, "exists the data.")
				End If
				Call InitialFieldValue(eventData)	
			Else		' update the data
				Call UpdatePatientInfo(eventData)
				MyStatusMSG(eventData, "successed updating the data.")
			End If
		Else
			MyStatusMSG(eventData, "should to input nom, prenom and dateofbirth.")
		End If
		Call GetPatientList(eventData)
	End Sub
	
	Sub deleteBtn_OnChange(ByVal eventData As MFPEventData) 'TODO change <fieldName> to desired field name
		
		Dim id As String
		id = GetHiddenIDValue(eventData)
		
		If id <> "-1" Then		
			Call DeletePatientInfo(eventData, id)
			Call MyStatusMSG(eventData, "successed deleting the data.")	
		End If
		'Call InitialFieldValue(eventData)	
		Call InitialFieldValue(eventData)
		Call GetPatientList(eventData)
	End Sub
	
	Function InitialFieldValue(ByVal eventData As MFPEventData)
		Call SetHiddenIDValue(eventData, "-1")
		eventData.Form.Fields.GetField("nom").Value = ""
		eventData.Form.Fields.GetField("prenom").Value = ""
		eventData.Form.Fields.GetField("dateofbirth").Value = ""
	End Function
	
	Function CreatePatientInfo(ByVal eventData As MFPEventData)
		Dim url As String = "https://stasindev.sharepoint.com/sites/Patient/_api/Web/Lists/getbytitle('patient')/Items"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "POST"
		request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		
		Dim nom As String = eventData.Form.Fields.GetField("nom").Value
		Dim prenom As String = eventData.Form.Fields.GetField("prenom").Value
		Dim dateofbirth As String = eventData.Form.Fields.GetField("dateofbirth").Value
		
		Dim json_data As String = "{""__metadata"":{""type"":""SP.Data.PatientListItem""},"
		json_data = json_data & """nom"": """ & nom & ""","
		json_data = json_data & """prenom"": """ & prenom & ""","
		json_data = json_data & """dateofbirth"": """ & dateofbirth & """}"
		
		Dim json_bytes() As Byte = System.Text.Encoding.ASCII.GetBytes(json_data)
		request.ContentLength = json_bytes.Length
		
    
		
		Dim stream As IO.Stream = request.GetRequestStream

		stream.Write(json_bytes, 0, json_bytes.Length)


		Dim response As HttpWebResponse = request.GetResponse
		Dim statusCode As String = response.StatusCode.ToString
		
		Dim dataStream As IO.Stream = response.GetResponseStream()
		Dim reader As New IO.StreamReader(dataStream)          ' Open the stream using a StreamReader for easy access.
		Dim responseFromServer As String = reader.ReadToEnd()  ' Read the content.

		' Cleanup the streams and the response.
		reader.Close()
		dataStream.Close()
		response.Close()
		CreatePatientInfo = statusCode
	End Function
	
	Function UpdatePatientInfo(ByVal eventData As MFPEventData)
		Dim id As String
		id = GetHiddenIDValue(eventData)
		Dim url As String = "https://stasindev.sharepoint.com/sites/Patient/_api/Web/Lists/getbytitle('patient')/Items/getbyid('" & id & "')"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "POST"
		request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		request.Headers.Add("If-Match", "*")
		request.Headers.Add("X-HTTP-Method", "MERGE")
		
		Dim nom As String = eventData.Form.Fields.GetField("nom").Value
		Dim prenom As String = eventData.Form.Fields.GetField("prenom").Value
		Dim dateofbirth As String = eventData.Form.Fields.GetField("dateofbirth").Value
		
		Dim json_data As String = "{""__metadata"":{""type"":""SP.Data.PatientListItem""},"
		json_data = json_data & """nom"": """ & nom & ""","
		json_data = json_data & """prenom"": """ & prenom & ""","
		json_data = json_data & """dateofbirth"": """ & dateofbirth & """}"
		
		Dim json_bytes() As Byte = System.Text.Encoding.ASCII.GetBytes(json_data)
		request.ContentLength = json_bytes.Length
		
    
		
		Dim stream As IO.Stream = request.GetRequestStream

		stream.Write(json_bytes, 0, json_bytes.Length)


		Dim response As HttpWebResponse = request.GetResponse
		Dim statusCode As String = response.StatusCode.ToString
		
		Dim dataStream As IO.Stream = response.GetResponseStream()
		Dim reader As New IO.StreamReader(dataStream)          ' Open the stream using a StreamReader for easy access.
		Dim responseFromServer As String = reader.ReadToEnd()  ' Read the content.

		' Cleanup the streams and the response.
		reader.Close()
		dataStream.Close()
		response.Close()
		UpdatePatientInfo = statusCode
	End Function
	
	Function DeletePatientInfo(ByVal eventData As MFPEventData, ByVal id As String)
		Dim url As String = "https://stasindev.sharepoint.com/sites/Patient/_api/Web/Lists/getbytitle('patient')/Items/getbyid('" & id & "')"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "DELETE"
		request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		request.Headers.Add("If-Match", "*")
		
		Dim response As HttpWebResponse = request.GetResponse
		Dim statusCode As String = response.StatusCode.ToString
		
		response.Close()
		DeletePatientInfo = statusCode
	End Function
	
	Function SetHiddenIDValue(ByVal eventData As MFPEventData, ByVal id As String)
		Dim curID As LabelField = eventData.Form.Fields.GetField("curID")
		curID.Text = id
		curID.IsHidden = True
	End Function
	
	Function GetHiddenIDValue(ByVal eventData As MFPEventData)
		Dim curID As LabelField = eventData.Form.Fields.GetField("curID")
		GetHiddenIDValue = curID.Text		
	End Function
	
	Function MyStatusMSG(ByVal eventData As MFPEventData, ByVal strMSG As String)
		Dim status As LabelField = eventData.Form.Fields.GetField("status")
		status.Text = strMSG
		status.IsHidden = False		
	End Function
	
	Function GetPaitentInfoByID(ByVal eventData As MFPEventData, ByVal id As String)
		Dim url As String = "https://stasindev.sharepoint.com/sites/Patient/_api/Web/Lists/getbytitle('patient')/Items(" & id & ")"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "GET"
		request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		
		If Not response Is Nothing Then response.Close()
		
		
		json = Mid(json, 6, json.Length - 6)
		json = "[" & json & "]"
		
	
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		
		Dim FieldNameArray() As String = FIELD_STRING.Split("|")
		Dim ItemNameArray() As String = ITEM_STRING.Split("|")
		Dim field As TextField
		
		For index As Integer = 0 To results.Count - 1
			results1 = results.Item(index)			
			For j As Integer = 0 To ItemNameArray.Length - 1
				
				If ItemNameArray(j) = "dateofbirth" Then
					
					Dim dateofbirth As DateField = eventData.Form.Fields.GetField(FieldNameArray(j))
					Dim iDate As String = results1.Item(ItemNameArray(j))
					Dim oDate As DateTime = Convert.ToDateTime(iDate)
					dateofbirth.Value = oDate.Date
					
				Else
					
					field = eventData.Form.Fields.GetField(FieldNameArray(j))				
					field.Value = results1.Item(ItemNameArray(j))
					
				End If				
				
			Next			
		Next		
		
		GetPaitentInfoByID = response.StatusCode
	End Function
	Function GetPatientList(ByVal eventData As MFPEventData)
		Dim url As String = "https://stasindev.sharepoint.com/sites/Patient/_api/Web/Lists/getbytitle('patient')/Items"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "GET"
		request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		
		If Not response Is Nothing Then response.Close()
		
		json = Mid(json, 17, json.Length - 18)
		
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		Dim documentList As ListField = eventData.Form.Fields.GetField("patientList")

		documentList.FindMode = False
		documentList.Items.Clear()
		documentList.MaxSearchResults = results.Count + 1
		
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		Dim listItem As ListItem 
		documentList.MaxSearchResults = results.Count + 1

		For index As Integer = 0 To results.Count - 1
			results1 = results.Item(index)			
			listItem = New ListItem(results1.Item("ID") & "| " & results1.Item("nom") & " " & results1.Item("prenom"), results1.Item("ID"))
			documentList.Items.Add(listItem)
			
		Next	
		
		GetPatientList = response.StatusCode
	End Function
	Function IsExistPatientInfo(ByVal eventData As MFPEventData, ByVal searchName As String)
		Dim url As String = "https://stasindev.sharepoint.com/sites/Patient/_api/Web/Lists/getbytitle('patient')/Items"
		
		Dim address As Uri = New Uri(url)
	
		Dim request As HttpWebRequest = DirectCast(WebRequest.Create(address), HttpWebRequest)
	
		request.Method = "GET"
		request.ContentType = "application/json;odata=verbose"
		request.Accept = "application/json;odata=verbose"
		request.Headers.Add("Authorization", "Bearer " & SECURITY_TOKEN)
		
		Dim response As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
		
		Dim reader As StreamReader = New StreamReader(response.GetResponseStream())
		
		Dim json As String = reader.ReadToEnd()
		
		If Not response Is Nothing Then response.Close()
		
		json = Mid(json, 17, json.Length - 18)
		
		Dim serializer As JavaScriptSerializer = New JavaScriptSerializer()
		
		Dim results As  System.Collections.ArrayList = serializer.Deserialize(Of System.Collections.ArrayList)(json)
		
		Dim documentList As ListField = eventData.Form.Fields.GetField("patientList")

		
		Dim results1 As  System.Collections.Generic.Dictionary(Of String, Object) 
		Dim listItem As ListItem 
		
		Dim patientName As String
		Dim bFound As Boolean
		bFound = False
		For index As Integer = 0 To results.Count - 1
			results1 = results.Item(index)			
			listItem = New ListItem(results1.Item("ID") & "| " & results1.Item("nom") & " " & results1.Item("prenom"), results1.Item("ID"))
			documentList.Items.Add(listItem)
			patientName = results1.Item("nom") & " " & results1.Item("prenom")
			If searchName = patientName Then
				bFound = True
				Exit For
			End If
			
		Next	
		
		IsExistPatientInfo = bFound
	End Function
End Module
