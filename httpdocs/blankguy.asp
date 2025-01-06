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
    
<div class="main-container">
    <!-- Page Heading -->
    <div class="page-heading">I'm a ____ guy</div>

    <div class="content-layout">
        <!-- Sidebar -->
        <div class="sidebar">
            <button id="random-entry-btn" class="action-button">Random Entry</button>
        </div>

        <!-- Main Content -->
        <div class="main-content">
            <%
            ' CREATE DATABASE CONNECTION
            Set cn = Server.CreateObject("ADODB.Connection")
            strConnect = Application("JOM_ConnectionString") ' Ensure this connection string is properly configured
            cn.Open(strConnect)

            ' SQL STATEMENT TO FETCH DATA FROM BlankGuy TABLE
            Set rs = cn.Execute("SELECT * FROM BlankGuy")

            ' CHECK IF THERE ARE RECORDS TO DISPLAY
            If Not rs.EOF Then
            %>
                <table class="data-table">
                    <caption>List of quotes from Blank Guy episodes.</caption>
                    <tr>
                        <th scope="col">ID</th>
                        <th scope="col">Quote</th>
                        <th scope="col">Episode #</th>
                        <th scope="col">Time</th>
                        <th scope="col">Audio</th>
                    </tr>
                    <%
                    ' LOOP THROUGH THE RECORDS
                    Do Until rs.EOF
                    %>
                    <tr>
                        <td><%= rs("ID") %></td>
                        <td><%= rs("Quote") %></td>
                        <td><%= rs("Ep#") %></td>
                        <td><%= rs("Time") %></td>
                        <td>
                            <% If rs("AudioFile") <> "" Then %>
                                <audio controls>
                                    <source src="/Audiofiles/<%= rs("AudioFile") %>" type="audio/mpeg">
                                    Your browser does not support the audio element.
                                </audio>
                            <% Else %>
                                No audio available
                            <% End If %>
                        </td>
                    </tr>
                    <%
                        rs.MoveNext
                    Loop
                    %>
                </table>
            <%
            Else
                Response.Write("<p class='error-message'>No records found.</p>")
            End If

            ' CLOSE THE DATABASE CONNECTION
            rs.Close
            cn.Close
            %>
        </div>
    </div>

    <!-- Floating Buttons -->
    <div class="floating-buttons">
        <button id="random-entry-btn-floating" class="action-button">Random Entry</button>
        <button id="back-to-top" onclick="scrollToTop()" class="action-button back-to-top">Back to Top</button>
    </div>
</div>

<!-- Styles -->
<style>
    .main-container {
        max-width: 1200px;
        margin: 20px auto;
        padding: 20px;
        background-color: #f9f9f9;
        border-radius: 8px;
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        font-family: Arial, Helvetica, sans-serif;
    }

    .page-heading {
        font-size: 2.5em;
        font-weight: bold;
        color: black;
        text-align: center;
        margin: 20px 0;
    }

    .content-layout {
        display: flex;
        gap: 20px;
    }

    .sidebar {
        width: 200px;
        background-color: #f4f4f4;
        padding: 10px;
        border-right: 1px solid #ddd;
    }

    .action-button {
        display: block;
        width: 100%;
        padding: 10px;
        background-color: #000;
        color: #fff;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        text-align: center;
        font-size: 1.1em;
        margin-bottom: 20px;
    }

    .action-button:hover {
        background-color: #333;
    }

    .floating-buttons {
        position: fixed;
        bottom: 20px;
        left: 20px;
        display: none;
        flex-direction: column;
        gap: 10px;
    }

    .back-to-top {
        width: auto;
    }

    .main-content {
        flex: 1;
    }

    .data-table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 20px;
        background-color: white;
    }

    .data-table th, .data-table td {
        padding: 10px;
        border: 1px solid #ddd;
        text-align: center;
    }

    .data-table th {
        background-color: #f4f4f4;
        font-weight: bold;
    }

    .data-table td audio {
        display: block;
        margin: 0 auto;
    }

    .error-message {
        text-align: center;
        color: red;
        margin: 20px 0;
        font-size: 1.2em;
    }

    @media (max-width: 768px) {
        .content-layout {
            flex-direction: column;
        }

        .sidebar {
            width: 100%;
            border-right: none;
            border-bottom: 1px solid #ddd;
            margin-bottom: 20px;
        }

        .data-table th, .data-table td {
            display: block;
            width: 100%;
            text-align: left;
        }

        .data-table th {
            background-color: #fff;
            font-size: 1.1em;
        }

        .data-table td {
            padding: 8px 5px;
        }
    }
</style>

<!-- JavaScript -->
<script>
    function scrollToTop() {
        window.scrollTo({ top: 0, behavior: "smooth" });
    }

    function goToRandomEntry() {
        const entries = document.querySelectorAll('.data-table tr');
        const randomIndex = Math.floor(Math.random() * (entries.length - 1)) + 1; // Avoid header row
        const randomEntry = entries[randomIndex];
        if (randomEntry) {
            randomEntry.scrollIntoView({ behavior: "smooth" });
            randomEntry.style.backgroundColor = "#ffeb3b"; // Highlight the row temporarily
            setTimeout(() => randomEntry.style.backgroundColor = "", 2000);
        }
    }

    function toggleFloatingButtons() {
        const floatingButtons = document.querySelector('.floating-buttons');
        if (window.scrollY > 200) {
            floatingButtons.style.display = 'flex';
        } else {
            floatingButtons.style.display = 'none';
        }
    }

    window.addEventListener('scroll', toggleFloatingButtons);

    document.getElementById("random-entry-btn").addEventListener("click", goToRandomEntry);
    document.getElementById("random-entry-btn-floating").addEventListener("click", goToRandomEntry);
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
