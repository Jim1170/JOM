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
    
<div class="main-container" style="display: flex; flex-direction: column; max-width: 1200px; margin: 0 auto;">
    <!-- Page Heading -->
    <div class="page-heading" style="text-align: center; font-size: 2.5em; font-weight: bold; color: #000; margin: 20px 0; width: 100%;">Dictionary</div>

    <!-- Content Layout -->
    <div style="display: flex; min-height: 100vh;">
        <!-- Sidebar -->
        <div class="sidebar" style="width: 200px; background-color: #f4f4f4; padding: 10px; border-right: 1px solid #ddd; overflow-y: auto;">
            <!-- Random Entry Button -->
            <button id="random-entry-btn" style="width: 100%; padding: 10px; margin-bottom: 20px; font-size: 1em; cursor: pointer; background-color: #000; color: #fff; border: none; border-radius: 5px;">Go to a Random Entry</button>

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
            <!-- Entries Container -->
            <div class="dictionary-entries" style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px;">
                <% 
                On Error Resume Next

                ' Create database connection
                Set cn = Server.CreateObject("ADODB.Connection")
                cn.Open(Application("JOM_ConnectionString"))

                If Err.Number <> 0 Then
                    Response.Write("<p style='color: red;'>Database connection error: " & Err.Description & "</p>")
                    Response.End
                End If

                ' Fetch and order entries
                sqlQuery = "SELECT Word, Pronunciation, Partofspeech, Definition, Examplesentence, ID, Category, Show FROM Dictionary ORDER BY Word ASC"
                Set rs = cn.Execute(sqlQuery)

                Dim currentLetter, isNumberSection
                currentLetter = ""
                isNumberSection = False

                If Not rs.EOF Then
                    Do While Not rs.EOF
                        Dim firstLetter
                        firstLetter = Left(rs("Word"), 1)

                        ' Group numeric entries under #
                        If IsNumeric(firstLetter) Then
                            If Not isNumberSection Then
                                isNumberSection = True
                                Response.Write("<h2 id='Numbers' style='grid-column: span 2; font-size: 1.8em; margin: 20px 0 10px; border-bottom: 2px solid #ddd;'>#</h2>")
                            End If
                        ElseIf firstLetter <> currentLetter Then
                            currentLetter = firstLetter
                            Response.Write("<h2 id='" & currentLetter & "' style='grid-column: span 2; font-size: 1.8em; margin: 20px 0 10px; border-bottom: 2px solid #ddd;'>" & currentLetter & "</h2>")
                        End If

                        ' Display entry with word as heading
                        Response.Write("<div class='entry'>")
                        Response.Write("<p style='font-weight: bold; font-size: 1.5em; margin-bottom: 5px;'>" & rs("Word") & "</p>")
                        Response.Write("<div style='border: 1px solid #ddd; padding: 15px; background-color: #fff; border-radius: 5px;'>")
                        If Not IsNull(rs("Pronunciation")) Then
                            Response.Write("<p><strong>Pronunciation:</strong> " & rs("Pronunciation") & "</p>")
                        End If
                        If Not IsNull(rs("Partofspeech")) Then
                            Response.Write("<p><strong>Part of Speech:</strong> " & rs("Partofspeech") & "</p>")
                        End If
                        If Not IsNull(rs("Category")) Then
                            Response.Write("<p><strong>Category:</strong> " & rs("Category") & "</p>")
                        End If
                        If Not IsNull(rs("Definition")) Then
                            Response.Write("<p><strong>Definition:</strong> " & rs("Definition") & "</p>")
                        End If
                        If Not IsNull(rs("Examplesentence")) Then
                            Response.Write("<p><strong>Example Sentence:</strong> " & rs("Examplesentence") & "</p>")
                        End If
                        If Not IsNull(rs("Show")) Then
                            Response.Write("<p><strong>Show:</strong> " & rs("Show") & "</p>")
                        End If
                        Response.Write("</div>")
                        Response.Write("</div>")

                        rs.MoveNext
                    Loop
                Else
                    Response.Write("<p>No entries found.</p>")
                End If

                rs.Close
                cn.Close
                %>
            </div>
        </div>
    </div>
</div>

<!-- Back to Top Button -->
<button id="back-to-top" onclick="scrollToTop()" style="position: fixed; bottom: 20px; left: 220px; padding: 10px 15px; font-size: 1em; cursor: pointer; background-color: #000; color: #fff; border: none; border-radius: 5px; display: none; z-index: 1000;">Back to Top</button>

<script>
    function scrollToTop() {
        window.scrollTo({ top: 0, behavior: "smooth" });
    }

    document.getElementById("random-entry-btn").addEventListener("click", () => {
        const totalEntries = document.querySelectorAll('.entry').length;
        const randomIndex = Math.floor(Math.random() * totalEntries);
        const randomEntry = document.querySelectorAll('.entry')[randomIndex];
        if (randomEntry) {
            randomEntry.scrollIntoView({ behavior: "smooth" });
            randomEntry.style.backgroundColor = "#fffbcc";
            setTimeout(() => (randomEntry.style.backgroundColor = ""), 2000);
        }
    });

    window.addEventListener("scroll", () => {
        const backToTopButton = document.getElementById("back-to-top");
        backToTopButton.style.display = window.scrollY > 200 ? "block" : "none";
    });
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
