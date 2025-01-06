<!-- send_email.asp -->
<%
'RETREIVE DATA FROM FORM
formName=(Request("Name"))
formEmail=(Request("Email"))
formSubject=(Request("Subject"))
formComment=(Request("message"))


'GET IP ADDRESS
Dim UserIPAddress
UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserIPAddress = "" Then
UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
End If


CommentOK="yes"
If InStr(formComment,"http")>0 Then
CommentOK="no"
End If

'ID DATA AVAILABLE SEND EMAIL
If NOT(IsEmpty(formEmail) ) AND CommentOK="yes" Then


'sTo = "mark@tmtech.com.au"
sTo = "james@regulationlore.com.au"

sFrom = "James@regulationlore.com.au"
sSubject = "Regulationlore Contact Us"
sBody = "Regulationlore have sent the following: Name:  " &  formName & vbCrLf  &  "  Email: " & formEmail & vbCrLf & "  Comment: " & formComment & vbCrLf & "  Subject: " & formSubject & vbCrLf & "  IP Address: " & UserIPAddress 


'sBCC="mark@tmtech.com.au"
sMailServer = "127.0.0.1"

Sub SendEMail()

	Set objMail = Server.CreateObject("CDO.Message")
	Set objConf = Server.CreateObject("CDO.Configuration") 
	Set objFields = objConf.Fields
	
	With objFields
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") ="mail.smtp2go.com"
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") ="mark@tmtech.com.au"		
		.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="mark2468"
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
		.Update 
	End With

	With objMail
		Set .Configuration = objConf
		.From = sFrom
		.To = sTo
		.BCC=sBCC
		.Subject = sSubject
		.TextBody = sBody
	End With
    
    objMail.Send
	
	Set objFields = Nothing
	Set objConf = Nothing
	Set objMail = Nothing
End sub

SendEMail()	

'Response.Redirect("Thanks.asp")


 If Err.Number = 0 Then
        Response.Write("Thank you for your message. We will get back to you soon.")
    Else
        Response.Write("Sorry, there was an error sending your message. Please try again later.")
    End If

End If


%>
