<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>



<meta content="en-au" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>regulation lore *</title>
<style type="text/css">
.auto-style1 {
	text-align: center;
}
.auto-style2 {
	font-family: Arial, Helvetica, sans-serif;
	font-size: xx-large;
	color: #FFFFFF;
}
.auto-style3 {
	border-width: 0;
	background-color: #000000;
}
.auto-style6 {
	border-width: 0;
	text-align: center;
	color: #FFFFFF;
	font-family: Arial, Helvetica, sans-serif;
	background-color: #000000;
}
.auto-style7 {
	color: #FFFFFF;
}
</style>

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
        .details {
            display: none;
            margin-left: 20px;
            background-color: #f0f0f0;
            padding: 10px;
            border-radius: 5px;
        }
        .word {
            cursor: pointer;
            font-weight: bold;
        }
    .auto-style8 {
	text-align: center;
}
    .auto-style9 {
	color: #FFFFFF;
	border: 1px solid #000000;
	background-color: #000000;
}
    </style>
    <script language="javascript" type="text/javascript">
        function toggleDetails(elementId) {
            // Close any currently open details
            var allDetails = document.getElementsByClassName('details');
            for (var i = 0; i < allDetails.length; i++) {
                if (allDetails[i].id !== elementId) {
                    allDetails[i].style.display = 'none';
                }
            }

            // Toggle the clicked details section
            var details = document.getElementById(elementId);
            if (details.style.display === "block") {
                details.style.display = "none";
            } else {
                details.style.display = "block";
            }
        }

        // Close details if clicking outside of them
        document.addEventListener('click', function(event) {
            var detailsElements = document.getElementsByClassName('details');
            var isClickInside = false;

            for (var i = 0; i < detailsElements.length; i++) {
                if (detailsElements[i].contains(event.target)) {
                    isClickInside = true;
                    break;
                }
            }

            if (!isClickInside && !event.target.classList.contains('word')) {
                for (var i = 0; i < detailsElements.length; i++) {
                    detailsElements[i].style.display = 'none';
                }
            }
        });
    </script>


</head>

<body>

<table style="width: 100%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="auto-style3" colspan="5"><a href="index.html"><span class="auto-style2">
		<strong><span class="auto-style7">regulation</span></strong></span><strong><br class="auto-style2" />
		<span class="auto-style2">lore *</span></a><span class="auto-style2"><br />
		</span></strong></td>
	</tr>
	<tr>
		<td class="auto-style6" style="width: 20%; height: 50px;"><strong>
		<a href="dictionary.html"><span class="auto-style7">Dictionary</span></a></strong></td>
		<td class="auto-style6" style="width: 20%; height: 50px;">
		<a href="blankguy.html"><span class="auto-style7"><strong>I'm a ____ guy</strong></span></a></td>
		<td class="auto-style6" style="width: 20%; height: 50px;">
		<strong><a href="regulationideas.html"><span class="auto-style7">
		Regulation Ideas</span></a></strong></td>
		<td class="auto-style6" style="width: 20%; height: 50px;"><strong>
		<a href="todo.html"><span class="auto-style7">To Do</span></a></strong></td>
		<td class="auto-style6" style="width: 20%; height: 50px;">
		<a href="collectables.html"><span class="auto-style7"><strong>Regulation 
		Collectables</strong></span></a></td>
	</tr>
	<tr>
		<td colspan="5"><br />
    

 <form method="post" action="">
 	<div class="auto-style8">
 	   Enter Search Text and Click Search
 <input name="SearchText" type="text" />
 <input name="Submit1" type="submit" value="submit" />
 </div>
 </form>

    

    <table style="width: 30%" class="auto-style5">
        <tr>
            <td style="width: 10%" class="auto-style9"><strong>Word</strong></td>
        </tr>
    </table>

    <%
    
    'GET VARIABLES FROM FORM
	SearchText=(Request("SearchText"))

        
    RecordCount = 0

    'CREATE DATABASE CONNECTION
    set cn = server.createobject("adodb.connection")
    strConnect = Application("JOM_ConnectionString")
    cn.open(strConnect)

    'SQL STATEMENT TO SEARCH
    'set rs = cn.Execute("SELECT * FROM Dictionary Order by word " )
    'set rs = cn.Execute("SELECT * FROM Dictionary WHERE Word like '" &  SearchText & "'" )   
    set rs = cn.Execute("SELECT * FROM Dictionary WHERE Word like '%" &  SearchText & "%'" )  
  
    
    while not rs.eof
        ' Generate a unique ID for the details div
        uniqueId = "details-" & rs("ID")
    %>

    <div class="word-container">
        <div class="word" onclick="toggleDetails('<%= uniqueId %>')">
            <% Response.Write rs("word") %>
        </div>
        <div class="details" id="<%= uniqueId %>">
            <p><strong>Pronunciation:</strong> <% Response.Write rs("Pronunciation") %></p>
            <p><strong>Part of Speech:</strong> <% Response.Write rs("Partofspeech") %></p>
            <p><strong>Definition:</strong> <% Response.Write rs("Definition") %></p>
            <p><strong>Example Sentence:</strong> <% Response.Write rs("Examplesentence") %></p>
        </div>
    </div>

    <%
        RecordCount = RecordCount + 1
        rs.movenext
    wend
    rs.close
    cn.close
    %>

    <p align="center">&nbsp;</p>
	</tr>
	<tr>
		<td class="auto-style1" style="width: 20%"><a href="contactus.asp">
		Contact Us</a></td>
		<td class="auto-style1" style="width: 20%">
		<a href="https://regulationstore.com/">Official Regulation Store</a></td>
		<td style="width: 20%">&nbsp;</td>
		<td style="width: 20%">&nbsp;</td>
		<td style="width: 20%">&nbsp;</td>
	</tr>
	<tr>
		<td class="auto-style1" colspan="5">This is a fan made site. </td>
	</tr>
</table>

</body>

</html>
