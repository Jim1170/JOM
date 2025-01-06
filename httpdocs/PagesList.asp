<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="en-au" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Pages</title>
<style type="text/css">
.auto-style1 {
	text-align: center;
}
.auto-style2 {
	border: 1px solid #000000;
}
</style>
<% 
Session.LCID = 3081
'If NOT(Request.Cookies("AccountType")="Admin") Then Response.Redirect("LogOut.asp") 
%>

</head>

<body>

<table style="width: 50%" align="center">
	<tr>
		<td class="auto-style2" style="width: 15%"><strong>Container </strong></td>
		<td class="auto-style2" style="width: 15%"><strong>Page Range</strong></td>

	</tr>
</table>
<%
'CREATE DATABSE CONNECTION
set cn=server.createobject("adodb.connection")
strConnect=Application("Container_ConnectionString")
cn.open(strConnect)

'SQL STATEMENT TO SEARCH 
'set rs=cn.Execute("SELECT Max(CIDNo) AS MaxPage FROM CData")
'set rs=cn.Execute("SELECT Max(PageNo) AS MaxPage FROM CData  WHERE ContainerNo='11' GROUP BY PageNo")
'set rs=cn.Execute("SELECT Judge, SUM(TotalTime) AS SUMTT FROM JudgeTime WHERE HearingDate < #1 Jan 2015# GROUP BY Judge")
'set rs=cn.Execute("SELECT ContainerNo, First(PageNo)AS MinPage, Last(PageNo) AS MaxPage FROM CData  GROUP BY ContainerNo")


set rs=cn.Execute("SELECT ContainerNo, First(PageNo)AS MinPage, Last(PageNo) AS MaxPage FROM CData  GROUP BY ContainerNo")

while not rs.eof
%>
<table style="width: 50%" align="center">
	<tr>
		<td class="auto-style2" style="width: 15%"><strong><% Response.write(rs("ContainerNo")) %> </strong></td>
		<td class="auto-style2" style="width: 15%"><strong><% Response.write(rs("MinPage") & "-" & rs("MaxPage")) %></strong></td>

	</tr>
</table>

<%
rs.movenext
wend
rs.close
cn.close

%>



<p class="auto-style1">List Pages for Container</p>
<p class="auto-style1"><a href="../AdminHome.asp">Admin Home</a></p>

</body>

</html>
