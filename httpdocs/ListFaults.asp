<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
.auto-style1 {
	font-size: medium;
}
.auto-style3 {
	border: 1px solid #000000;
}
.auto-style5 {
	border-color: #000000;
	border-width: 0;
	border-collapse: collapse;
}
.auto-style6 {
	text-align: center;
}
.auto-style7 {
	font-size: x-small;
}
.auto-style8 {
	border: 1px solid #000000;
	font-size: x-small;
	font-family: Arial, Helvetica, sans-serif;
}
.auto-style9 {
	font-size: x-small;
	font-family: Arial, Helvetica, sans-serif;
}
</style>

<% 
Session.LCID = 3081

'CHECK TO SEE THEY HAVE COOKIE SET FROM LOGIN
TestUser=Request.Cookies("ValidUser")
If Not (TestUser="IT" ) Then Response.Redirect ("Default.asp")
%>


</head>

<body>
<p align="center" class="auto-style1"><strong>List Fault</strong><p align="center">
<span class="auto-style9">&nbsp; 
</span> 
<a href="Home.asp"><span class="auto-style9">Home</span></a></p>
<form action="" method="post">
	<div class="auto-style6" style="height: 70px">
		<select name="FaultStatus" class="auto-style9">
		<option selected="" value="All">All</option>
		<option value="Completed">Completed</option>
		<option value="Current">Current</option>
		</select><br class="auto-style9">
		<input name="btnSubmit" type="Submit" value="Submit" class="auto-style9"><br class="auto-style9">
	</div>
</form>
<p align="center" class="auto-style9">&nbsp;</p>


<table style="width: 100%" class="auto-style5">
	<tr>
		<td style="width: 5%" class="auto-style8"><strong>Fault No</strong></td>
		<td style="width: 10%" class="auto-style8"><strong>ID/HWare</strong></td>
		<td style="width: 6%" class="auto-style8"><strong>Date Open</strong></td>
		<td style="width: 8%" class="auto-style8"><strong>Date Close</strong></td>
		<td style="width: 8%" class="auto-style8"><strong>Location</strong></td>
		<td style="width: 8%" class="auto-style8"><strong>Priority</strong></td>
		<td style="width: 5%" class="auto-style8"><strong>Lost Audio</strong></td>
		<td style="width: 6%" class="auto-style8"><strong>Result</strong></td>
		<td style="width: 13%" class="auto-style8"><strong>Description</strong></td>
		<td style="width: 7%" class="auto-style8"><strong>Engineer</strong></td>
		<td class="auto-style8" style="width: 8%"><strong>Status</strong></td>
	</tr>
</table>
<%
'IS SUBMIT IS CLICKED
If Request.Form("btnSubmit") = "Submit"  Then 

'GET VARIABLES FROM FORM
FaultStatus=(Request("FaultStatus"))


set cn=server.createobject("adodb.connection")
strConnect=Application("CRSHW_ConnectionString")
cn.open(strConnect)

If FaultStatus="All" Then
'SQL STATEMENT TO SEARCH BY ALL FAULTS
set rs=cn.Execute("SELECT * FROM Fault WHERE Location like '" & Request.Cookies("UserLocation") & "'")

Else

'SQL STATEMENT TO SEARCH BY FAULT STATUS
set rs=cn.Execute("SELECT * FROM Fault WHERE FaultStatus like '" & FaultStatus & "' AND Location like '" & Request.Cookies("UserLocation") & "'")

End if

while not rs.eof

%>

<table style="width: 100%" class="auto-style5">
	<tr>
		<td style="width: 5%" class="auto-style8">
		<a href="UpdateFault.asp?FaultID=<%Response.Write rs("FaultID")%>"><%Response.Write rs("FaultID")%></a></td>
		<td style="width: 10%" class="auto-style3">
		<a href="update.asp?IDNo=<%Response.Write rs("IDNo")%>">
		<span class="auto-style9"><%Response.Write rs("IDNo")%></span></a>/<span class="auto-style9"><%Response.Write rs("Hardware")%></span></td>
		<td style="width: 6%" class="auto-style8"><%Response.Write rs("DateofFault")%></td>
		<td style="width: 8%" class="auto-style8"><%Response.Write rs("DateClosed")%></td>
		<td style="width: 8%" class="auto-style8"><%Response.Write rs("Location")%></td>
		<td style="width: 8%" class="auto-style8"><%Response.Write rs("Priority")%></td>
		<td style="width: 5%" class="auto-style3"><span class="auto-style9"><%Response.Write rs("PrimaryAudioLost")%>
		</span>/<span class="auto-style9"><%Response.Write rs("BackupAudioLost")%></span></td>
		<td style="width: 6%" class="auto-style8"><%Response.Write rs("Result")%></td>
		<td style="width: 13%" class="auto-style8"><%Response.Write rs("Description")%></td>
		<td style="width: 7%" class="auto-style8"><%Response.Write rs("Engineer")%></td>
		<td class="auto-style8" style="width: 8%"><%Response.Write rs("FaultStatus")%></td>
	</tr>
</table>





<%

rs.movenext
wend
rs.close
cn.close

End If

%></table><p align="center" class="auto-style9">&nbsp;</p>
<p align="center" class="auto-style7">&nbsp;</p>
</body>

</html>


