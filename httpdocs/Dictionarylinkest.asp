<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Regulation Lore</title>
    <style>
        body {
            margin: 0;
            font-family: Arial, Helvetica, sans-serif;
            background-color: #ffffff;
        }

        table {
            width: 100%;
            border-collapse: collapse;
        }

        .header {
            background-color: #000000;
            color: #ffffff;
        }

        .header img {
            float: left;
            margin: 10px;
        }

        .header h1 {
            font-size: 50pt;
            margin: 0;
            display: inline-block;
            vertical-align: middle;
            line-height: 115px;
        }

        .header p {
            font-size: xx-large;
            margin: 0;
            color: #ffffff;
        }

        .navbar {
            background-color: #000000;
            text-align: center;
        }

        .navbar a {
            color: #ffffff;
            text-decoration: none;
            padding: 10px 20px;
            display: inline-block;
        }

        .navbar a:hover {
            text-decoration: underline;
        }

        .footer {
            text-align: center;
            font-size: small;
            margin: 20px 0;
        }

        .footer img {
            margin: 10px;
        }

        .dynamic-text {
            color: #333333;
            font-size: 10pt;
        }
    .auto-style1 {
		border-width: 0px;
	}
    </style>
</head>

<body>

    <!-- Header Section -->
    <table class="header">
        <tr>
            <td>
                <a href="index.asp">
                    <img src="Images/FanMadeWhite.png" alt="FanMade Logo" height="145" width="101" class="auto-style1" /></a><a href="index2.html">
                </a>
                <h1>Regulation Lore *</h1>
                <p>A fan-made site</p>
            </td>
        </tr>
    </table>

    <!-- Navigation Bar -->
    <table class="navbar">
        <tr>
            <td><a href="dictionary.asp">Dictionary</a></td>
            <td><a href="blankguy.asp">I'm a ____ guy</a></td>
            <td><a href="castideas.asp">Regulation Ideas</a></td>
            <td><a href="todo.asp">To Do</a></td>
            <td><a href="RegulationCollectables.asp">Regulation Collectables</a></td>
        </tr>
    </table>

    <!-- Placeholder for Page Content -->
    
<!-- Page Heading -->
<div class="page-heading">Dictionary</div>

<!-- Dictionary Viewer -->
<div class="container" style="margin: 20px;">
    <% 
    On Error Resume Next ' Gracefully handle errors

    ' CREATE DATABASE CONNECTION
    Set cn = Server.CreateObject("ADODB.Connection")
    strConnect = Application("JOM_ConnectionString")
    cn.Open(strConnect)
    If Err.Number <> 0 Then
        Response.Write("<p style='color: red; text-align: center;'>Error connecting to the database: " & Err.Description & "</p>")
        Response.End
    End If

    ' Function to sanitize inputs and prevent SQL injection
    Function FixSQL(str)
        If IsNull(str) Or str = "" Then
            FixSQL = ""
        Else
            FixSQL = Replace(str, "'", "''")
        End If
    End Function

    ' Function to create links in definitions
    Function CreateLinks(definition)
        Dim regEx, matches, match, relatedWord, relatedWordAnchor
        ' Regular expression to find "See 'XXXX'"
        Set regEx = New RegExp
        regEx.Pattern = "See\s'([^']+)'"
        regEx.Global = True
        ' Find matches
        Set matches = regEx.Execute(definition)
        ' Replace matches with HTML links
        For Each match In matches
            relatedWord = match.SubMatches(0) ' Extract the word inside quotes
            ' Generate the anchor link for the related word
            relatedWordAnchor = Replace(Server.HTMLEncode(relatedWord), " ", "_")
            ' Replace "See 'XXXX'" with an HTML link that triggers navigateToEntry
            definition = Replace(definition, match.Value, "<a href=""#"" onclick=""navigateToEntry('" & relatedWordAnchor & "'); return false;"">See '" & relatedWord & "'</a>")
        Next
        CreateLinks = definition
    End Function

    ' Check if Clear Filters was clicked
    If Request("ClearFilters") <> "" Then
        SearchText = ""
        PartOfSpeechFilter = ""
        CategoryFilter = ""
        ShowFilter = ""
    Else
        ' Retrieve filter variables
        SearchText = Request("SearchText")
        PartOfSpeechFilter = Request("PartOfSpeechFilter")
        CategoryFilter = Request("CategoryFilter")
        ShowFilter = Request("ShowFilter")
    End If

    ' Build SQL query
    sqlQuery = "SELECT * FROM Dictionary WHERE 1=1"
    If SearchText <> "" Then
        sqlQuery = sqlQuery & " AND (Word LIKE '%" & FixSQL(SearchText) & "%'" & _
                               " OR Pronunciation LIKE '%" & FixSQL(SearchText) & "%'" & _
                               " OR Partofspeech LIKE '%" & FixSQL(SearchText) & "%'" & _
                               " OR Definition LIKE '%" & FixSQL(SearchText) & "%'" & _
                               " OR Examplesentence LIKE '%" & FixSQL(SearchText) & "%')"
    End If
    If PartOfSpeechFilter <> "" Then
        sqlQuery = sqlQuery & " AND Partofspeech = '" & FixSQL(PartOfSpeechFilter) & "'"
    End If
    If CategoryFilter <> "" Then
        sqlQuery = sqlQuery & " AND Category = '" & FixSQL(CategoryFilter) & "'"
    End If
    If ShowFilter <> "" Then
        sqlQuery = sqlQuery & " AND Show = '" & FixSQL(ShowFilter) & "'"
    End If

    ' Add ORDER BY clause to ensure alphabetical sorting
    sqlQuery = sqlQuery & " ORDER BY Word ASC"

    ' Execute query
    Set rs = cn.Execute(sqlQuery)
    If Err.Number <> 0 Then
        Response.Write("<p style='color: red; text-align: center;'>Error executing SQL query: " & Err.Description & "</p>")
        Response.End
    End If
    %>

    <!-- Search and Filters -->
    <div style="text-align: center; margin-bottom: 20px;">
        <form method="post" action="">
            <!-- Search Bar -->
            <label for="searchText">Search:</label>
            <input id="searchText" name="SearchText" type="text" value="<%= Server.HTMLEncode(SearchText) %>" />
            <input name="SubmitSearch" type="submit" value="Search" />

            <!-- Filters -->
            <label for="partOfSpeechFilter">Part of Speech:</label>
            <select id="partOfSpeechFilter" name="PartOfSpeechFilter" onchange="this.form.submit()">
                <option value="">All</option>
                <%
                Set rsPos = cn.Execute("SELECT DISTINCT Partofspeech FROM Dictionary ORDER BY Partofspeech")
                If Not rsPos Is Nothing And Not rsPos.EOF Then
                    Do While Not rsPos.EOF
                        posValue = rsPos("Partofspeech")
                %>
                <option value="<%= posValue %>" <% If posValue = PartOfSpeechFilter Then Response.Write "selected" %>><%= Server.HTMLEncode(posValue) %></option>
                <%
                        rsPos.MoveNext
                    Loop
                End If
                If Not rsPos Is Nothing Then rsPos.Close
                %>
            </select>

            <label for="categoryFilter">Category:</label>
            <select id="categoryFilter" name="CategoryFilter" onchange="this.form.submit()">
                <option value="">All</option>
                <%
                Set rsCategory = cn.Execute("SELECT DISTINCT Category FROM Dictionary ORDER BY Category")
                If Not rsCategory Is Nothing And Not rsCategory.EOF Then
                    Do While Not rsCategory.EOF
                        categoryValue = rsCategory("Category")
                %>
                <option value="<%= categoryValue %>" <% If categoryValue = CategoryFilter Then Response.Write "selected" %>><%= Server.HTMLEncode(categoryValue) %></option>
                <%
                        rsCategory.MoveNext
                    Loop
                End If
                If Not rsCategory Is Nothing Then rsCategory.Close
                %>
            </select>

            <label for="showFilter">Show:</label>
            <select id="showFilter" name="ShowFilter" onchange="this.form.submit()">
                <option value="">All</option>
                <%
                Set rsShow = cn.Execute("SELECT DISTINCT Show FROM Dictionary ORDER BY Show")
                If Not rsShow Is Nothing And Not rsShow.EOF Then
                    Do While Not rsShow.EOF
                        showValue = rsShow("Show")
                %>
                <option value="<%= showValue %>" <% If showValue = ShowFilter Then Response.Write "selected" %>><%= Server.HTMLEncode(showValue) %></option>
                <%
                        rsShow.MoveNext
                    Loop
                End If
                If Not rsShow Is Nothing Then rsShow.Close
                %>
            </select>

            <!-- Clear Button -->
            <input name="ClearFilters" type="submit" value="Clear Filters" />
        </form>
    </div>

    <!-- Word Results -->
    <table style="width: 90%; border-collapse: collapse; margin: 20px auto;">
        <thead>
            <tr>
                <th style="background-color: black; color: white; padding: 10px; text-align: left;">Word</th>
            </tr>
        </thead>
        <tbody>
            <% If Not rs.EOF Then %>
            <% Do While Not rs.EOF %>
            <tr>
                <td style="border: 1px solid black; padding: 10px;">
                    <!-- Add unique ID to the word container for anchor navigation -->
                    <div id="<%= Replace(Server.HTMLEncode(rs("Word")), " ", "_") %>" class="word" onclick="toggleDetails('<%= rs("ID") %>')" style="font-weight: bold; cursor: pointer;">
                        <%= Server.HTMLEncode(rs("Word")) %>
                    </div>
                    <!-- Add a unique ID for the details section -->
                    <div class="details" id="<%= rs("ID") %>" style="margin-top: 10px; display: none;">
                        <p><strong>Pronunciation:</strong> <%= Server.HTMLEncode(rs("Pronunciation")) %></p>
                        <p><strong>Part of Speech:</strong> <%= Server.HTMLEncode(rs("Partofspeech")) %></p>
                        <p><strong>Category:</strong> <%= Server.HTMLEncode(rs("Category")) %></p>
                        <p><strong>Definition:</strong> <%= CreateLinks(rs("Definition")) %></p>
                        <p><strong>Example Sentence:</strong> <%= Server.HTMLEncode(rs("Examplesentence")) %></p>
                    </div>
                </td>
            </tr>
            <% rs.MoveNext %>
            <% Loop %>
            <% Else %>
            <tr>
                <td style="text-align: center; color: red;">No results found.</td>
            </tr>
            <% End If %>
        </tbody>
    </table>

    <% 
    ' Close database objects
    If Not rs Is Nothing Then rs.Close
    If Not cn Is Nothing Then cn.Close
    %>
</div>

<!-- JavaScript for Toggle and Smooth Scrolling -->
<script>
    // Toggle visibility of details for a specific word
    function toggleDetails(id) {
        const details = document.getElementById(id); // The details div to toggle
        const isVisible = details.style.display === "block";

        // Collapse all details
        document.querySelectorAll(".details").forEach(el => el.style.display = "none");

        // Expand the clicked entry if it wasn't already visible
        if (!isVisible) {
            details.style.display = "block";
        }
    }

    // Navigate to and expand the specified entry
    function navigateToEntry(targetId) {
        // Collapse all entries
        document.querySelectorAll(".details").forEach(el => el.style.display = "none");

        // Scroll to the target entry
        const targetElement = document.getElementById(targetId);
        if (targetElement) {
            targetElement.scrollIntoView({ behavior: "smooth" });

            // Expand the target entry
            const detailsElement = targetElement.nextElementSibling; // Get the associated .details div
            if (detailsElement && detailsElement.classList.contains("details")) {
                detailsElement.style.display = "block";
            }
        }
    }

    // Handle navigation to a specific entry on page load
    document.addEventListener("DOMContentLoaded", function () {
        const targetId = window.location.hash.substring(1); // Extract the hash (e.g., #entry-id)
        if (targetId) {
            navigateToEntry(targetId); // Navigate to and expand the target entry
        }
    });
</script>

<!-- Smooth Scroll CSS -->
<style>
    html {
        scroll-behavior: smooth;
    }
    .details {
        display: none;
    }
</style>

    <!-- Footer Section -->
    <table class="footer">
        <tr>
            <td colspan="5">
                <p class="dynamic-text">
                    <a href="About.asp">About</a> |
                    <a href="contactus.asp">Comment Leavers</a> |
                    <a href="https://regulationstore.com/" target="_blank" rel="noopener noreferrer">Official Regulation Store</a>
                </p>
				<p class="dynamic-text">
                    This is a fan-made site brought to you by UniFan. (<span id="dynamic-text"></span>)
                </p>
                <img src="Images/UniFan.png" alt="UniFan Logo" height="67" width="94" /></td>
        </tr>
    </table>

    <!-- Add JavaScript for Dynamic Text -->
    <script>
        document.addEventListener('DOMContentLoaded', () => {
            const textOptions = [
                "Unifan, combining the power of one with the fabric of the fan. Unifan",
                "Unifan, combining the fabric of the fan with the power of one. Unifan"
            ];
            const randomText = textOptions[Math.floor(Math.random() * textOptions.length)];
            document.getElementById("dynamic-text").textContent = randomText;
        });
    </script>

	<p>
	<a href="https://www.youtube.com/watch?v=qy_ui7uU00s" rel="noopener noreferrer" target="_blank">
	<img alt="" class="auto-style1" height="65" src="Images/Mousehole.png" style="float: right" width="99"></a><a href="Gurpcountdown.asp"><img alt="" class="auto-style1" height="68" src="Images/Gurp.png" style="float: right" width="48"></a><a href="Bleep.asp"><img alt="" class="auto-style1" height="57" src="Images/Bleep.png" style="float: right" width="86"></a></p>

</body>

</html>
