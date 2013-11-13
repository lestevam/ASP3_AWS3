<script language="javascript" runat="server">
	function GMTNow(){return new Date().toGMTString()}
</script>
<%
CONST AWS_BUCKETNAME = "name_bucketname"
CONST AWS_ACCESSKEY = "your_accesskey"
CONST AWS_SECRETKEY = "your_secretkey"

class AWS3

	private AWSBucketUrl
	private strNow
	private fileStream
	
	Private Sub class_initialize()
		AWSBucketUrl = "http://s3.amazonaws.com/" & AWS_BUCKETNAME	
		Set fileStream = CreateObject("ADODB.Stream")
		strNow = GMTNow()
	End Sub
	
	Dim Signature
	Dim StringToSign
	Dim Authorization
	
	Public Sub deleteFolderAWS(source)
		StringToSign = Replace("DELETE\n\n\n" & strNow & "\n/" & AWS_BUCKETNAME & "/" & source, "\n", vbLf)
		Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))
		Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature	
		
		With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			.open "DELETE", AWSBucketUrl & "/" & source, False
			.setRequestHeader "Authorization", Authorization
			.setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com"  
			.setRequestHeader "Date",strNow
			.send
			If .status <> 200 Then 
				Response.ContentType = "text/xml"
				Response.Write .responseText
			End If
		End With
		
	End Sub
	
	Public Sub deleteFilesAWS(source)
		StringToSign = Replace("DELETE\n\n\n" & strNow & "\n/" & AWS_BUCKETNAME & "/" & source, "\n", vbLf)
		Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))
		Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature	
		
		With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			.open "DELETE", AWSBucketUrl & "/" & source, False
			.setRequestHeader "Authorization", Authorization
			.setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com"  
			.setRequestHeader "Date",strNow
			.send
			If .status <> 200 Then 
				Response.ContentType = "text/xml"
				Response.Write .responseText
			End If
		End With
		
	End Sub
	
	Private Sub downloadAWS(source)
		StringToSign = Replace("GET\n\n\n" & strNow & "\n/" & AWS_BUCKETNAME & "/" & source, "\n", vbLf)
		Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))
		Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature	
		
		With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			.open "GET", AWSBucketUrl & "/" & source, False
			.setRequestHeader "Authorization", Authorization
			.setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com"  
			.setRequestHeader "Date",strNow
			.send
			If .status = 200 Then
				Set oStream = CreateObject("ADODB.Stream")
				oStream.Open
				oStream.Type = 1
				oStream.Write .responseBody
				oStream.Position = 0
				Set fileStream = oStream
			Else 
				Response.ContentType = "text/xml"
				Response.Write .responseText
			End If
		End With
		
	End Sub
	
	Private Sub uploadAWS(destination)
		StringToSign = Replace("PUT\n\n\n" & strNow & "\nx-amz-acl:public-read\n/" & AWS_BUCKETNAME & "/" & destination, "\n", vbLf)
		Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))
		Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature	
		
		With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			.open "PUT", AWSBucketUrl & "/" & destination, False
			.setRequestHeader "Authorization", Authorization
			.setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com"  
			.setRequestHeader "Date",strNow
			.setRequestHeader "x-amz-acl", "public-read"
			.send fileStream.Read
			If .status <> 200 Then ' successful
				Response.ContentType = "text/xml"
				Response.Write .responseText
			End If
			fileStream.close
		End With
		
	End Sub	
	
	Function moveFiles(source, destination)
		strNow = GMTNow()
		StringToSign = Replace("GET\n\n\n" & strNow & "\n/" & AWS_BUCKETNAME & "/", "\n", vbLf)
		Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))
		Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature	
		
		With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			.open "GET", AWSBucketUrl & "/?prefix=" & source, False
			.setRequestHeader "Authorization", Authorization
			.setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com" 
			.setRequestHeader "Date",strNow
			.send
			If .status = 200 Then
				Dim objxml
				Set objxml = Server.CreateObject("MSXML2.FreeThreadedDOMDocument")
				objxml.async = False
				objxml.setProperty "SelectionLanguage", "XPath"
				objxml.loadXML .responseText
				
				Set xmlList = objxml.getElementsByTagName("Key")
				For Each xmlItem in xmlList
					Dim fileSource : fileSource = xmlItem.childNodes(0).text
					if ( InStr(fileSource,".") > 0 ) Then
						Dim fileDestination : fileDestination = Replace(fileSource,source,destination)
						
						Call downloadAWS(fileSource)
						Call uploadAWS(fileDestination)
						Call deleteFilesAWS(fileSource)					
					End If
				next
			Else 
				Response.ContentType = "text/xml"
				Response.Write .responseText
			End If
		End With
		
	End Function
	
	Private Sub Download(source, destination)
		destination = Server.Mappath(destination)
		StringToSign = Replace("GET\n\n\n" & strNow & "\n/" & AWS_BUCKETNAME & source, "\n", vbLf)
		Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))
		Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature	
		
		With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			.open "GET", AWSBucketUrl & source, False
			.setRequestHeader "Authorization", Authorization
			.setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com"  
			.setRequestHeader "Date",strNow
			.send
			If .status = 200 Then
				Set oStream = CreateObject("ADODB.Stream")
				oStream.Open
				oStream.Type = 1
				oStream.Write .responseBody
				oStream.SaveToFile destination
				oStream.Close
			Else 
				Response.ContentType = "text/xml"
				Response.Write .responseText
			End If
		End With
		
	End Sub
	
	Private Sub upload(source, destination)
		source = Server.Mappath(source)
		StringToSign = Replace("PUT\n\n\n" & strNow & "\nx-amz-acl:public-read\n/" & AWS_BUCKETNAME & destination, "\n", vbLf)
		Signature = BytesToBase64(HMACSHA1(AWS_SECRETKEY, StringToSign))
		Authorization = "AWS " & AWS_ACCESSKEY & ":" & Signature	
		
		With Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			.open "PUT", AWSBucketUrl & destination, False
			.setRequestHeader "Authorization", Authorization
			'.setRequestHeader "Content-Type", "text/plain"
			.setRequestHeader "Host", AWS_BUCKETNAME & ".s3.amazonaws.com"  
			.setRequestHeader "Date",strNow
			.setRequestHeader "x-amz-acl", "public-read"
			.send GetBytes(source) 
			If .status = 200 Then 
				Response.Write "<a href="& AWSBucketUrl & sRemoteFilePath &" target=_blank>Uploaded File</a>"
			Else
				Response.ContentType = "text/xml"
				Response.Write .responseText
			End If
		End With
		
	End Sub

	Private Function GetBytes(sPath)
		With Server.CreateObject("Adodb.Stream")
			.Type = 1 ' adTypeBinary
			.Open
			.LoadFromFile sPath
			.Position = 0
			GetBytes = .Read
			.Close
		End With
	End Function

	Private Function BytesToBase64(varBytes)
		With Server.CreateObject("MSXML2.DomDocument").CreateElement("b64")
			.dataType = "bin.base64"
			.nodeTypedValue = varBytes
			BytesToBase64 = .Text
		End With
	End Function

	Private Function HMACSHA1(varKey, varValue)
		With Server.CreateObject("System.Security.Cryptography.HMACSHA1")
			.Key = UTF8Bytes(varKey)
			HMACSHA1 = .ComputeHash_2(UTF8Bytes(varValue))
		End With
	End Function

	Private Function UTF8Bytes(varStr)
		With Server.CreateObject("System.Text.UTF8Encoding")
			UTF8Bytes = .GetBytes_4(varStr)
		End With
	End Function
	
end class
%>