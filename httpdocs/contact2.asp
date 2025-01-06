<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Untitled 1</title>
</head>

<body>


<form method="POST" >
         <input TYPE="hidden" NAME="VTI-GROUP" VALUE="0">
            <table border="0" cellpadding="2" cellspacing="0" width="600">
              <tr>
                <td width="27"></td>
                <td width="176"></td>
                <td width="549"></td>
              </tr>
              <tr>
                <td width="27"></td>
                <td width="176"><font face="Century Gothic" size="3">Name </font></td>
                <td width="549"><input type="text" name="formName" size="26"></td> 
              </tr>
              <tr>
                <td width="27"></td>
                <td width="176"><font face="Century Gothic" size="3">Email&nbsp;</font></td>
                <td width="549"><input type="text" name="formEmail" size="37"></td>
              </tr>
              <tr>
                <td width="27">&nbsp;</td>
                <td width="176"><font face="Century Gothic">Phone (optional)</font></td>
                <td width="549"><input type="text" name="formPhone" size="37"></td>
              </tr>
              <tr>
                <td width="27"></td>
                <td width="176"><font face="Century Gothic" size="3">Comment</font><font face="Century Gothic">
                </font>
                  <p>&nbsp;</p>
                </td>
                <td width="549">
				<textarea rows="7" name="formComment" cols="50"></textarea></td>
              </tr>
            </table>
          <p align="center"><font size="4" face="Century Gothic">
		  <input type="submit" value="Submit" name="B3"><input type="reset" value="Reset" name="B4"></font></p>
        </form>
        <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><o:p>
        				<div class="auto-style2">
							<span style="color:#333333" class="auto-style1">
        <font color="#808080"><strong></div>
        				<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
        				&nbsp;<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
        				&nbsp;<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
        				&nbsp;<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
        				&nbsp;</td>
					</tr>
				</table>
        <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><span style="font-size:10.0pt;font-family:&quot;Century Gothic&quot;;color:#333333"><o:p>
        </td>
    </tr>
    </table>
</div>

<%
'RETREIVE DATA FROM FORM
formName=(Request("formName"))
formEmail=(Request("formEmail"))
formComment=(Request("formComment"))
formPhone=(Request("formPhone"))

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


sTo = "mark@tmtech.com.au"

sFrom = "mark@tmtech.com.au"
sSubject = "Regulation"
sBody = "RegulationLore Contact Form has sent the following: Name:  " &  formName & vbCrLf  &  "  Email: " & formEmail & vbCrLf & "  Comment: " & formComment & vbCrLf & "  Phone: " & formPhone & vbCrLf & "  IP Address: " & UserIPAddress 


sBCC="mark@tmtech.com.au"
'sMailServer = "127.0.0.1"

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





</body>

</html>
