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
<p align="center" class="auto-style1">Dictionary<p align="center">&nbsp; 
<a href="Index.html">Home</a></p>
<p align="center"></p>


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

<a href="Test2.asp?formWord=<%Response.Write rs("word")%>"><%Response.Write rs("word")%></a>
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


