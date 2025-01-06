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
</style>

<% 
Session.LCID = 3081

%>

</head>

<body>



<p align="center" class="auto-style1">Dictionary- Test 5<p align="center">&nbsp; 
<a href="Index.html">Home</a></p>
<p align="center"></p>

<a href="../html-link.htm" target="popup" onclick="window.open('../html-link.htm','name','width=600,height=400')">Open page in new window</a>

<table style="width: 30%" class="auto-style5">
	<tr>
		<td style="width: 10%" class="auto-style3"><strong>Word</strong></td>
		
	</tr>
</table>
<%
RecordCount=0

'CREATE DATABASE CONNECTION
set cn=server.createobject("adodb.connection")
strConnect=Application("JOM_ConnectionString")
cn.open(strConnect)

'SQL STATEMENT TO SEARCH
set rs=cn.Execute("SELECT * FROM Dictionary" ) 
while not rs.eof

 
%>

<table style="width: 30%" class="auto-style5">
	<tr>
		<td style="width: 10%" class="auto-style3">
		
<a href="Test2.asp?formWord=<%Response.Write rs("word")%>"target="popup"  onclick="window.open('Test2.asp?formWord=<%Response.Write rs("word")%>','name','width=600,height=400')"><%Response.Write rs("word")%></a></td>
					
						
	</tr>
</table>

<%

RecordCount=RecordCount+1

rs.movenext
wend
rs.close
cn.close

%>

<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
</body>

</html>


