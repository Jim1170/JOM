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

        header, nav, footer {
            width: 100%;
        }

        header {
            background-color: #000000;
            color: #ffffff;
            display: flex;
            align-items: center;
            padding: 10px;
        }

        header img {
            height: 100px;
            margin-right: 15px;
        }

        header h1 {
            font-size: 2.5em;
            margin: 0;
        }

        header p {
            margin: 0;
            font-size: 1.2em;
            color: #ffffff;
        }

        nav {
            background-color: #000000;
            display: flex;
            justify-content: center;
            padding: 10px 0;
        }

        nav ul {
            list-style: none;
            margin: 0;
            padding: 0;
            display: flex;
            gap: 20px;
        }

        nav a {
            color: #ffffff;
            text-decoration: none;
            padding: 10px 15px;
        }

        nav a:hover {
            text-decoration: underline;
        }

        footer {
            text-align: center;
            padding: 20px;
            background-color: #f1f1f1;
            font-size: 0.9em;
            position: relative;
        }

        footer a {
            color: #000000;
            text-decoration: none;
        }

        footer a:hover {
            text-decoration: underline;
        }

        footer img {
            width: 80px; /* Adjusted for slightly larger size */
            height: auto; /* Maintain aspect ratio */        }

        .dynamic-text {
            color: #333333;
        }

        .easter-eggs {
            position: fixed; /* Anchors to the bottom-right corner */
            bottom: 10px;
            right: 10px;
            display: flex;
            gap: 10px; /* Spacing between images */
        }

        .easter-eggs img {
            width: 40px; /* Smaller size for easter eggs */
            height: auto; /* Maintain aspect ratio */
            cursor: pointer;
            transition: transform 0.2s; /* Hover effect */
        }

        .easter-eggs img:hover {
            transform: scale(1.1); /* Slight zoom on hover */
        }
    .auto-style1 {
		font-weight: bold;
	}
    </style>
</head>

<body>
    <!-- Header Section -->
    <header>
        <a href="index.asp">
            <img src="Images/FanMadeWhite.png" alt="FanMade Logo" />
        </a>
        <div>
            <h1>Regulation Lore *</h1>
            <p>A fan-made site</p>
        </div>
    </header>

    <!-- Navigation Bar -->
    <nav>
        <ul>
            <li><a href="dictionary.asp">Dictionary</a></li>
            <li><a href="blankguy.asp">I'm a ____ guy</a></li>
            <li><a href="castideas.asp">Regulation Ideas</a></li>
            <li><a href="todo.asp">To Do</a></li>
            <li><a href="RegulationCollectables.asp">Regulation Collectables</a></li>
        </ul>
    </nav>

    <!-- Placeholder for Page Content -->
    
<div class="main-container" style="display: flex; flex-direction: column; min-height: 100vh; max-width: 1200px; margin: 0 auto;">
    <% 
    ' Initialize database connection
    Dim cn, rs, rsPos, rsCategory, rsShow, sqlQuery, searchText, partOfSpeechFilter, categoryFilter, showFilter, totalEntries
    On Error Resume Next
    Set cn = Server.CreateObject("ADODB.Connection")
    cn.Open(Application("JOM_ConnectionString"))

    If Err.Number <> 0 Then
        Response.Write("<p style='color: red;'>Database connection error: " & Err.Description & "</p>")
        Response.End
    End If

    ' Handle search and filter inputs
    searchText = Trim(Request("SearchText"))
    partOfSpeechFilter = Trim(Request("PartOfSpeechFilter"))
    categoryFilter = Trim(Request("CategoryFilter"))
    showFilter = Trim(Request("ShowFilter"))

    ' Count total entries for Random Entry functionality
    Set rsCount = cn.Execute("SELECT COUNT(*) AS Total FROM Dictionary")
    If Not rsCount.EOF Then totalEntries = rsCount("Total") Else totalEntries = 0
    rsCount.Close

    ' Build SQL query
    sqlQuery = "SELECT Word, Pronunciation, Partofspeech, Definition, Examplesentence, ID, Category, Show FROM Dictionary WHERE 1=1"
    If searchText <> "" Then
        sqlQuery = sqlQuery & " AND (Word LIKE '%" & Replace(searchText, "'", "''") & "%' OR Definition LIKE '%" & Replace(searchText, "'", "''") & "%')"
    End If
    If partOfSpeechFilter <> "" Then
        sqlQuery = sqlQuery & " AND Partofspeech = '" & Replace(partOfSpeechFilter, "'", "''") & "'"
    End If
    If categoryFilter <> "" Then
        sqlQuery = sqlQuery & " AND Category = '" & Replace(categoryFilter, "'", "''") & "'"
    End If
    If showFilter <> "" Then
        sqlQuery = sqlQuery & " AND Show = '" & Replace(showFilter, "'", "''") & "'"
    End If
    sqlQuery = sqlQuery & " ORDER BY Word ASC"

    ' Fetch filter options for Part of Speech, Category, and Show
    Set rsPos = cn.Execute("SELECT DISTINCT Partofspeech FROM Dictionary WHERE Partofspeech IS NOT NULL AND Partofspeech <> '' ORDER BY Partofspeech")
    Set rsCategory = cn.Execute("SELECT DISTINCT Category FROM Dictionary WHERE Category IS NOT NULL AND Category <> '' ORDER BY Category")
    Set rsShow = cn.Execute("SELECT DISTINCT Show FROM Dictionary WHERE Show IS NOT NULL AND Show <> '' ORDER BY Show")
    %>

    <!-- Page Heading -->
    <div class="page-heading" style="text-align: center; font-size: 2.5em; font-weight: bold; color: #000; margin: 20px 0; width: 100%;">Dictionary</div>

    <!-- Content Layout -->
    <div style="display: flex; flex: 1;">
        <!-- Sidebar -->
        <div class="sidebar" style="width: 200px; background-color: #f4f4f4; padding: 10px; border-right: 1px solid #ddd; overflow-y: auto;">
            <!-- Random Entry Button -->
            <button id="random-entry-btn-sidebar" style="width: 100%; padding: 12px; margin-bottom: 20px; font-size: 1.1em; cursor: pointer; background-color: #000; color: #fff; border: none; border-radius: 5px;">Random Entry</button>

            <!-- Search Bar -->
            <form id="search-form" method="GET" action="" onsubmit="return preventJump();">
                <label for="searchText" style="font-weight: bold; display: block; margin-bottom: 5px;">Search:</label>
                <input type="text" id="searchText" name="SearchText" placeholder="Search..." style="width: 100%; padding: 8px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px;" value="<%= Server.HTMLEncode(searchText) %>">
                <button type="submit" style="width: 100%; padding: 8px; background-color: #000; color: #fff; border: none; border-radius: 4px; cursor: pointer;">Search</button>
            </form>

            <!-- Part of Speech Filter -->
            <form id="filter-form" method="GET" action="" onsubmit="return preventJump();">
                <label for="partOfSpeechFilter" style="font-weight: bold; display: block; margin-top: 20px; margin-bottom: 5px;">Part of Speech:</label>
                <select id="partOfSpeechFilter" name="PartOfSpeechFilter" style="width: 100%; padding: 8px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px;" onchange="this.form.submit()">
                    <option value="">All</option>
                    <% 
                    If Not rsPos.EOF Then
                        Do While Not rsPos.EOF
                            Dim posValue
                            posValue = rsPos("Partofspeech")
                            Response.Write("<option value='" & Server.HTMLEncode(posValue) & "'")
                            If posValue = partOfSpeechFilter Then Response.Write(" selected")
                            Response.Write(">" & Server.HTMLEncode(posValue) & "</option>")
                            rsPos.MoveNext
                        Loop
                    End If
                    rsPos.Close
                    %>
                </select>

                <!-- Category Filter -->
                <label for="categoryFilter" style="font-weight: bold; display: block; margin-top: 20px; margin-bottom: 5px;">Category:</label>
                <select id="categoryFilter" name="CategoryFilter" style="width: 100%; padding: 8px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px;" onchange="this.form.submit()">
                    <option value="">All</option>
                    <% 
                    If Not rsCategory.EOF Then
                        Do While Not rsCategory.EOF
                            Dim categoryValue
                            categoryValue = rsCategory("Category")
                            Response.Write("<option value='" & Server.HTMLEncode(categoryValue) & "'")
                            If categoryValue = categoryFilter Then Response.Write(" selected")
                            Response.Write(">" & Server.HTMLEncode(categoryValue) & "</option>")
                            rsCategory.MoveNext
                        Loop
                    End If
                    rsCategory.Close
                    %>
                </select>

                <!-- Show Filter -->
                <label for="showFilter" style="font-weight: bold; display: block; margin-top: 20px; margin-bottom: 5px;">Show:</label>
                <select id="showFilter" name="ShowFilter" style="width: 100%; padding: 8px; margin-bottom: 10px; border: 1px solid #ccc; border-radius: 4px;" onchange="this.form.submit()">
                    <option value="">All</option>
                    <% 
                    If Not rsShow.EOF Then
                        Do While Not rsShow.EOF
                            Dim showValue
                            showValue = rsShow("Show")
                            Response.Write("<option value='" & Server.HTMLEncode(showValue) & "'")
                            If showValue = showFilter Then Response.Write(" selected")
                            Response.Write(">" & Server.HTMLEncode(showValue) & "</option>")
                            rsShow.MoveNext
                        Loop
                    End If
                    rsShow.Close
                    %>
                </select>

                <!-- Clear Filters Button -->
                <button type="button" onclick="clearFilters()" style="width: 100%; padding: 8px; background-color: #d9534f; color: #fff; border: none; border-radius: 4px; cursor: pointer; margin-top: 10px;">Clear Filters</button>
            </form>

            <!-- A-Z Navigation -->
            <a href="#Numbers" style="display: block; margin-bottom: 8px; text-decoration: none; font-weight: bold; color: #333; padding: 5px 8px; background-color: #e0e0e0; border-radius: 4px; text-align: center;">#</a>
            <% 
            For i = 65 To 90
                Response.Write("<a href='#" & Chr(i) & "' style='display: block; margin-bottom: 8px; text-decoration: none; font-weight: bold; color: #333; padding: 5px 8px; background-color: #e0e0e0; border-radius: 4px; text-align: center;'>" & Chr(i) & "</a>")
            Next
            %>
        </div>

        <!-- Main Content -->
        <div class="main-content" style="flex: 1; padding: 20px;">
            <div class="dictionary-entries">
                <% 
                Dim currentLetter, firstLetter, isNumberSection
                currentLetter = ""
                isNumberSection = False

                Set rs = cn.Execute(sqlQuery)
                If Not rs.EOF Then
                    Do While Not rs.EOF
                        firstLetter = UCase(Left(rs("Word"), 1))

                        ' Group numeric entries under #
                        If IsNumeric(firstLetter) Then
                            If Not isNumberSection Then
                                isNumberSection = True
                                Response.Write("<h2 id='Numbers' style='font-size: 1.8em; margin: 20px 0 10px; border-bottom: 2px solid #ddd;'>#</h2>")
                            End If
                        ElseIf firstLetter <> currentLetter Then
                            currentLetter = firstLetter
                            Response.Write("<h2 id='" & currentLetter & "' style='font-size: 1.8em; margin: 20px 0 10px; border-bottom: 2px solid #ddd;'>" & currentLetter & "</h2>")
                        End If

                        ' Display entry with a subtle border
                        Response.Write("<div class='entry' id='entry-" & rs("ID") & "' style='border: 1px solid #ddd; margin-bottom: 10px; border-radius: 5px; overflow: hidden;'>")
                        Response.Write("<div class='accordion' onclick='toggleDetails(" & rs("ID") & ")' style='cursor: pointer; font-weight: bold; padding: 10px; background-color: #fff; color: #000;'>")
                        Response.Write("<span>" & rs("Word") & "</span> <span id='icon-" & rs("ID") & "' style='float: right;'>Show ▼</span>")
                        Response.Write("</div>")
                        Response.Write("<div class='details' id='details-" & rs("ID") & "' style='display: none; padding: 10px; background-color: #f9f9f9;'>")
                        Response.Write("<p><strong>Pronunciation:</strong> " & rs("Pronunciation") & "</p>")
                        Response.Write("<p><strong>Part of Speech:</strong> " & rs("Partofspeech") & "</p>")
                        Response.Write("<p><strong>Category:</strong> " & rs("Category") & "</p>")
                        Response.Write("<p><strong>Definition:</strong> " & rs("Definition") & "</p>")
                        Response.Write("<p><strong>Example Sentence:</strong> " & rs("Examplesentence") & "</p>")
                        Response.Write("<p><strong>Show:</strong> " & rs("Show") & "</p>")
                        Response.Write("</div></div>")
                        rs.MoveNext
                    Loop
                End If
                rs.Close
                cn.Close
                %>
            </div>
        </div>
    </div>
</div>

<!-- Back to Top and Random Entry Buttons -->
<button id="random-entry-btn-floating" style="position: fixed; bottom: 60px; left: 20px; padding: 15px; background-color: #000; color: #fff; border: none; border-radius: 5px; display: none; cursor: pointer; font-size: 1.1em;">Random Entry</button>
<button id="back-to-top" onclick="scrollToTop()" style="position: fixed; bottom: 20px; left: 20px; padding: 15px; background-color: #000; color: #fff; border: none; border-radius: 5px; display: none; cursor: pointer; font-size: 1.1em;">Back to Top</button>

<!-- JavaScript -->
<script>
    function scrollToTop() {
        window.scrollTo({ top: 0, behavior: "smooth" });
    }

    function toggleDetails(id) {
        const details = document.getElementById('details-' + id);
        const icon = document.getElementById('icon-' + id);
        const isVisible = details.style.display === "block";

        document.querySelectorAll('.details').forEach(el => el.style.display = "none");
        document.querySelectorAll('[id^="icon-"]').forEach(el => el.innerHTML = "Show ▼");

        if (!isVisible) {
            details.style.display = "block";
            icon.innerHTML = "Hide ▲";
        }
    }

    function goToRandomEntry() {
        const entries = document.querySelectorAll('.entry');
        const randomIndex = Math.floor(Math.random() * entries.length);
        const randomEntry = entries[randomIndex];
        if (randomEntry) {
            const rect = randomEntry.getBoundingClientRect();
            const offset = rect.top + window.scrollY - 100; // Adjust to avoid hiding
            window.scrollTo({ top: offset, behavior: "smooth" });
            toggleDetails(randomEntry.id.split('-')[1]); // Open accordion
        }
    }

    document.getElementById("random-entry-btn-sidebar").addEventListener("click", goToRandomEntry);
    document.getElementById("random-entry-btn-floating").addEventListener("click", goToRandomEntry);

    window.addEventListener("scroll", () => {
        const backToTopButton = document.getElementById("back-to-top");
        const randomButton = document.getElementById("random-entry-btn-floating");
        const show = window.scrollY > 200 ? "block" : "none";
        backToTopButton.style.display = show;
        randomButton.style.display = show;
    });

    function clearFilters() {
        const url = new URL(window.location.href);
        url.searchParams.delete("PartOfSpeechFilter");
        url.searchParams.delete("CategoryFilter");
        url.searchParams.delete("ShowFilter");
        window.location.href = url.toString();
    }

    function preventJump() {
        // Prevent form default behavior
        const searchParams = new URLSearchParams(window.location.search);
        searchParams.set('SearchText', document.getElementById("searchText").value);
        const categoryFilter = document.getElementById("categoryFilter");
        const partOfSpeechFilter = document.getElementById("partOfSpeechFilter");
        searchParams.set('CategoryFilter',categoryFilter.value);
        searchParams.set('PartOfSpeechFilter',partOfSpeechFilter.value);

        const actionUrl=window.location.pathname+'?'+searchParams;
        window.location.href=actionUrl;
        return false;

    }
</script>

    <!-- Footer Section -->
    <footer>
        <p class="dynamic-text">
            <a href="About.asp">About</a> |
            <a href="contactus.asp">Comment Leavers</a> |
            <a href="https://regulationstore.com/" target="_blank" rel="noopener noreferrer">Official Regulation Store</a>
        </p>
        <p class="dynamic-text">
            This is a fan-made site brought to you by UniFan. (<span id="dynamic-text"></span>)
        </p>
        <img src="Images/UniFan.png" alt="UniFan Logo" />

        <!-- Easter Egg Links -->
        <div class="easter-eggs">
            <a href="https://www.youtube.com/watch?v=qy_ui7uU00s" target="_blank" rel="noopener noreferrer">
                <img src="Images/Mousehole.png" alt="Mousehole Link" />
            </a>
            <a href="Gurpcountdown.asp">
                <img src="Images/Gurp.png" alt="Gurp Countdown" />
            </a>
            <a href="Bleep.asp">
                <img src="Images/Bleep.png" alt="Bleep Link" />
            </a>
        </div>
    </footer>

    <!-- Dynamic JavaScript -->
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
</body>

</html>
