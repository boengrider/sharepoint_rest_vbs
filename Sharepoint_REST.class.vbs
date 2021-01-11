Class SharePoint
	
	Private sClientID ' Provided by the user. You need to register an app on sharepoint
	Private sClientSecret ' Provided by the user. You need to register an app on sharepoint
	Private sResourceID ' Sometimes also called ClientID. Not to be confused with trusted app client id that we generated. This resourceid is needed in order to get security token 
	Private sSecurityToken ' Security token used for authentication. It's obtained from sAuthUrl1 + sTenantRealmID + sAuthUrl2
	Private sTenantRealmID ' Tenant/Realm ID
	Private sXRequestDigest ' aka FormDigestValue
	Private sSiteUrl ' Your site url
	Private sAuthUrl1 ' 1st part of the authentication url where we obtain the security token 
	Private sAuthUrl2 ' 2nd part of the authentication url
	Private sMSonlineUrl ' https://login.microsoftonline.com/[yourtenant].onmicrosoft.com/.well-known/openid-configuration
	Private sTenantName ' volvogroup e.g volvogroup.sharepoint.com -> Tenant name is volvogroup in this case
	Private numTokenValidFor ' Number of seconds token is valid for
	Private oHTTP ' MSXML2.XMLHTTP.6.0
	Private oHTTPSrv ' MSXML2.ServerXMLHTTP.6.0  Some calls fail (timeout) using XMLHTTP6.0 but they succeed using Server
	Private oRX ' Regular expression object
	
	Private Sub Class_Initialize
		
		Set oRX = New RegExp
		Set oHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
		Set oHTTPSrv = CreateObject("MSXML2.ServerXMLHTTP.6.0")
		sMSonlineUrl = "https://login.microsoftonline.com/[yourtenant].onmicrosoft.com/.well-known/openid-configuration"
		sAuthUrl1 = "https://accounts.accesscontrol.windows.net/"
		sAuthUrl2 = "/tokens/OAuth/2"
		sSiteUrl = Null 
		sTenantRealmID = Null
		sClientSecret = Null
		sSecurityToken = Null
		
	End Sub
	
	'******************** P u b l i c  &  P r i v a t e   M e t h o d s *************
	
	'******************** Init() **********************
	'Function initializes the object. Sets values and calls other methods
	'Arguments:
	'strSiteUrl -> your sharepoint site URL
	'strTenantName -> usually subdomain of your site. In my case volvogroup.sharepoint.com tenant name is volvogroup
	'strClientID -> client ID generated on your sharepoint via trusted app portal
	'strClientSecret -> client secrect generated on your sharepoint via trusted app portal
	'Return values:
	'0 -> All OK
	'1 -> Tenant ID not found
	'2 -> Resource ID not found
	'3 -> Both Tenant and Resource IDs not found
	'**************************************************
	
	Public Function Init(strSiteUrl,strTenantName,strClientID,strClientSecret)
	
		If Right(strSiteUrl,1) <> "/" Then
			sSiteUrl = strSiteUrl & "/"
		End If 
		sClientID = strClientID
		sClientSecret = strClientSecret
		sTenantName = strTenantName
		sMSonlineUrl = Replace(sMSonlineUrl,"[yourtenant]",sTenantName)
		GetTenantRealmID_vti
		
		Select Case GetTenantRealmID 
		
			Case 0 ' Both GUIDs are OK
				
				Select Case GetSecurityToken
				
					Case 0 ' OK
						
				End Select 
				
			Case 1 ' Tenant ID not found
				Init = 1
				Exit Function
				
			Case 2 ' Resource ID not found
				Init = 2
				Exit Function
				
			Case 3 ' Tenant and Resource ID not found
				Init = 3
				Exit Function
				
		End select
		
		Select Case GetXRequestDigest ' So far the last step in the initialization phase
		
			Case 0 ' XRequestDigest OK
				Init = 0
				Exit function
			Case 1 ' HTTP ok but no digest found
				Init = 4
				Exit Function
				
		End Select  
					
		
	End Function
	
	'****************** GetTenantRealmID() *******************
	' Function obtains two values needed for securing a security
	' token at later point:
	' Tenant/Realm ID 
	' Client/Resource ID
	' Function returns 0 if successfull, positive number otherwise
	'*************************************************************
	Private Function GetTenantRealmID
	
		Dim retval : retval = 0
		Dim colAVPs,avp,colMatches,temp
		On Error Resume Next 
		With oHTTPSrv
			.open "GET", sSiteUrl & "_vti_bin/client.svc", False 
			.setRequestHeader "Authorization", "Bearer"
			.send
		End With
		If oHTTPSrv.status = 401 Then 
			colAVPs = Split(oHTTPsrv.getResponseHeader("WWW-Authenticate"),",") ' We have attribute value pairs e.g. Bearer realm="f25493ae-1c98-41d7-8a33-0be75f5fe603",client_id="00000003-0000-0ff1-ce00-000000000000"
			For Each avp In colAVPs
				oRX.Pattern = "Bearer realm"
				oRX.IgnoreCase = True
				If oRX.Test(avp) Then ' Bearer realm found
					temp = Split(avp,"=")
					sTenantRealmID = Replace(temp(1),"""","")
					oRX.Pattern = "([a-fA-F0-9]{8}-){1}([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}" ' Bearer Realm aka Tenant ID guid
					If oRX.Test(sTenantRealmID) Then
						retval = retval Or 0 ' GUID ok
					Else 
						retval = retval Or 1 ' Error Tenant ID guid
					End If 
				End If
				
				oRX.Pattern = "client_id"
				oRX.IgnoreCase = True
				If oRX.Test(avp) Then
					temp = Split(avp,"=")
					sResourceID = Replace(temp(1),"""","")
					oRX.Pattern = "([a-fA-F0-9]{8}-){1}([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}" ' Resource ID guid
					If oRX.Test(sResourceID) Then
						retval = retval Or 0 ' GUID ok
					Else
						retval = retval Or 2 ' Error Resource ID GUID
					End If 
				End If 
			Next 
		End If 
		
		GetTenantRealmID_vti = retval ' Error of some kind
		
	End Function  
	
	'******************** GetSecurityToken() **********************
	' Function obtains a security token
	' Token is usually valid for 24 hours
	' Token is used for authroziation purposes
	'**************************************************************
	Private Function GetSecurityToken
	
		Dim part,colParts,tokens,token,strBody,colAVPs
		strBody = "grant_type=client_credentials&client_id=" & sClientID & "@" & sTenantRealmID & "&client_secret=" & sClientSecret & "&resource=" & sResourceID & "/volvogroup.sharepoint.com@" & sTenantRealmID
		With oHTTP
			.open "POST", sAuthUrl1 & sTenantID & sAuthUrl2, False
			.setRequestHeader "Host","accounts.accesscontrol.windows.net"
			.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			.setRequestHeader "Content-Length", CStr(Len(strBody))
			.send strBody
		End With 
		
		If Not oHTTP.status = 200 Then
			GetSecurityToken = oHTTP.status
			Exit Function 
		End If 
		
		oRX.IgnoreCase = True
		colParts = Split(oHTTP.responseText,",")
		For Each part In colParts
			oRX.Pattern = "expires_in"
			If oRX.Test(part) Then
				colAVPs = Split(part,":")
				numTokenValidFor = Replace(colAVPs(1),"""","")
			End If 
			oRX.Pattern = "access_token"
			If oRX.Test(part) Then
				colAVPs = Split(part,":")
				sSecurityToken = Replace(colAVPs(1),"""","")
				sSecurityToken = Replace(sSecurityToken,"}","")
			End If 
		Next 
		
		GetSecurityToken = 0 ' All OK
	End Function 
	
	'****************** GetXRequestDigest() ************************
	' Function obtains XReguestDigestValue aka FormDigestValue
	'***************************************************************
	Private Function GetXRequestDigest()
	
		Dim colMatches
		With oHTTP
			.open "POST", sSiteUrl & "_api/contextinfo", False 
			.setRequestHeader "Authorization", "Bearer " & sSecurityToken
			.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			.setRequestHeader "Accept", "atom+xml;odata=verbose"
			.send
		End With
		
		If oHTTP.status = 200 Then
			oRX.Pattern = "(<d:FormDigestValue>)[0-9]{1}[xX]{1}[a-fA-F0-9]+"
			oRX.IgnoreCase = True
			If oRX.Test(oHTTP.responseText) Then
				Set colMatches = oRX.Execute(oHTTP.responseText)
				oRX.Pattern = "[0-9]{1}[xX]{1}[a-fA-F0-9]+"
				Set colMatches = oRX.Execute(colMatches(0))
				sXRequestDigest = colMatches(0)
				GetXRequestDigest = 0
				Exit Function
			Else
				GetXRequestDigest = 1 ' HTTP ok but no digest found
				Exit Function
			End If 
		End If 
		
		
		
	End Function 
	
	'***********************************************************************
	'****************** Public methods for list manipulation ***************
	'***********************************************************************
	
	
	'***************** ListCreate() *******************
	' Function creates a new list if it doesn't exist
	' If the list already exists function overwrites 
	' the existing one, i.e. deletes the old one
	'**************************************************
	Public Function ListCreate(strListName,strListDescription)
		 
		 Dim strRequest : strRequest = "{""__metadata"": { ""type"": ""SP.List"" }, ""BaseTemplate"":""100"", ""Description"": """ & strListDescription & """, ""Title"":""" & strListName & """ }"
		 
		 With oHTTP
		 	.open "POST",sSiteUrl & "_api/web/lists", False
		 	.setRequestHeader "Authorization","Bearer " & sSecurityToken
		 	.setRequestHeader "Accept","application/atom+xml;odata=verbose"
		 	.setRequestHeader "Content-Type","application/json;odata=verbose"
		 	.setRequestHeader "X-RequestDigest",sXRequestDigest
		 	.setRequestHeader "Content-Length",Len(strRequest)
		 	.send strRequest
		 End With
		 
		 If oHTTP.status = 201 Or oHTTP.status = 200 Then
		 	ListCreate = 0
		 	Exit Function
		 End If 
		 
	End Function 
	
	
	
	
	'******************* P r o p e r t i e s ********************
	Public Property Get XRequestDigest
	
		XRequestDigest = sXRequestDigest
		
	End Property 
	
	Public Property Get TenantName
	
		TenantName = sTenantName
		
	End Property 
	
	Public Property Get LoginUrl
	
		LoginUrl = sMSonlineUrl
		
	End Property 
	
	Public Property Get TenantRealmID
	
		TenantRealmID = sTenantRealmID
		
	End Property 
	
	Public Property Get ResourceID
	
		ResourceID = sResourceID
		
	End Property 
	
	Public Property Get SiteUrl
	
		SiteUrl = sSiteUrl
		
	End Property 
	
	Public Property Get SecurityToken
	
		SecurityToken = sSecurityToken
		
	End Property
	
	Public Property Get TokenValidPeriod
	
		TokenValidPeriod = numTokenValidFor
		
	End Property 
	
End Class


Dim oSP : Set oSP = New SharePoint
' Initialize object:
'	1st argument -> Sharepoint site URL
'       2nd argument -> Tenant name or host part of the subdomain
'	3rd argument -> ClientID 
'	4th argument -> Client secret
oSP.Init "https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it","volvogroup","462ad7ed-XXXX-XXXX-b808-18c6f33fadd7","dWHEl4AMp8qHX/oFFFcY4RyFJJRD7z1cIavjDH53yIE="
' Create two lists:
' 	1st argument -> List name
' 	2nd argument -> List description
oSP.ListCreate "SharePointList#1","Created using SharepointClass"
oSP.ListCreate "SharePointList#2","Created using SharepointClass"


	
		
