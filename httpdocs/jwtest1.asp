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
    <p align="center" class="auto-style1">Dictionary - Test 5</p>
    <p align="center">&nbsp;<a href="Index.html">Home</a></p>

    <table style="width: 30%" class="auto-style5">
        <tr>
            <td style="width: 10%" class="auto-style3"><strong>Word</strong></td>
        </tr>
    </table>

    <%
    RecordCount = 0

    'CREATE DATABASE CONNECTION
    set cn = server.createobject("adodb.connection")
    strConnect = Application("JOM_ConnectionString")
    cn.open(strConnect)

    'SQL STATEMENT TO SEARCH
    set rs = cn.Execute("SELECT * FROM Dictionary")
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
</body>

</html>
