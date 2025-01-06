<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>View All Files</title>
<style type="text/css">
.auto-style1 {
	font-size: medium;
}
.auto-style5 {
	border-color: #000000;
	border-width: 0;
	border-collapse: collapse;
}
.auto-style6 {
	border: 1px solid #000000;
	text-align: center;
}
.auto-style8 {
	border: 1px solid #000000;
	font-size: small;
	font-family: Arial, Helvetica, sans-serif;
}
</style>

<% 
Session.LCID = 3081
'GET VARIABLES FROM FORM
formContainerNo=(Request("formContainerNo"))

%>



</head>

<body>
<p align="center" class="auto-style1"><strong>View All Files by Container</strong><p align="center" class="auto-style1">

<a href="Default.asp">Home</a></p>


<form method="post" action="">
<p align="center">Enter Container No:&nbsp;&nbsp;&nbsp; 
<input name="formContainerNo" type="text" value="<%=formContainerNo%>">&nbsp; </p>
	<p align="center"><input name="Submit1" type="submit" value="submit"></p>
<p align="center">&nbsp;</p>
</form>


<table style="width: 60%" class="auto-style5" align="center">
	<tr>
		<td style="width: 10%" class="auto-style6"><strong>
		<a href="ViewAllFiles.asp?formContainerNo=<%=formContainerNo%>">IDNo</a></strong></td>
		<td style="width: 20%" class="auto-style6"><strong>
		<a href="ViewAllFilesByCode.asp?formContainerNo=<%=formContainerNo%>">Code</a></strong></td>
		<td style="width: 30%" class="auto-style6"><strong>
		<a href="ViewAllFilesByFileNo.asp?formContainerNo=<%=formContainerNo%>">FileNo</a></strong></td>
		<td class="auto-style6" style="width: 20%"><strong>Container No</strong></td>
		<td class="auto-style6" style="width: 20%"><strong>Page No</strong></td>
	</tr>
</table>
<%

'GET VARIABLES FROM FORM
formContainerNo=(Request("formContainerNo"))


'IS THERE DATA TO INSERT
If NOT (formContainerNo="")   Then


'CREATE DATABSE CONNECTION
set cn=server.createobject("adodb.connection")
strConnect=Application("Container_ConnectionString")
cn.open(strConnect)

'SQL STATEMENT TO SEARCH BY FILE NUMBER

'set rs=cn.Execute("SELECT * FROM SData WHERE FileNo Like '" & formFileNo & "' AND FileCategory Like '" & FileCategory & "' AND FileYear Like '" & FileYear & "'")
set rs=cn.Execute("SELECT * FROM CData WHERE ContainerNo like '" & formContainerNo & "' ORDER by Code")

while not rs.eof

%>

<table style="width: 60%" class="auto-style5" align="center">
	<tr>
		<td style="width: 10%" class="auto-style8"><%Response.Write rs("CIDNo")%></td>
		<td style="width: 20%" class="auto-style8"><%Response.Write rs("Code")%></td>
		<td style="width: 30%" class="auto-style8"><%Response.Write rs("CaseNo")%></td>
		<td class="auto-style8" style="width: 20%"><%Response.Write rs("ContainerNo")%></td>
		<td class="auto-style8" style="width: 20%"><%Response.Write rs("PageNo")%></td>
		</tr>
</table>


<%
rs.movenext
wend
rs.close
cn.close

End If

%><p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
</body>

</html>


