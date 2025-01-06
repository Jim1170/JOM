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
    
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Welcome to Regulation Lore</title>
    <style>
        body {
            font-family: Arial, Helvetica, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f4f4f4;
            color: #333;
        }

        main {
            max-width: 1200px;
            margin: 20px auto;
            padding: 20px;
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
        }

        .intro {
            flex: 1;
            max-width: 50%;
            text-align: left;
            font-size: 1.2em;
            line-height: 1.6;
        }

        .entries-container {
            flex: 1 1 100%;
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            justify-content: space-between;
        }

        .entry-section {
            flex: 1;
            min-width: 300px;
            background-color: #fff;
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
        }

        .entry-title {
            font-size: 1.5em;
            font-weight: bold;
            margin-bottom: 10px;
            text-align: center;
        }

        .entry-content {
            font-size: 1em;
            line-height: 1.5;
        }

        audio {
            display: block;
            margin: 10px auto;
        }

        @media (max-width: 768px) {
            .intro {
                max-width: 100%;
                text-align: center;
            }

            .entries-container {
                flex-direction: column;
            }
        }
    </style>
</head>

<body>
    <main>
        <!-- Intro Section -->
        <div class="intro">
            <h1>Welcome to Regulation Lore</h1>
            <p>This is a fan-made site bringing you the deep lore behind the nothing.</p>
            <p>This site is still under construction and will continue to grow.</p>
        </div>

        <!-- Entries Section -->
        <div class="entries-container">
            <!-- Random Dictionary Entry -->
            <section class="entry-section">
                <div class="entry-title">Today's Random Dictionary Entry</div>
                <div class="entry-content">
                    <% 
                    ' Create database connection
                    Set cn = Server.CreateObject("ADODB.Connection")
                    strConnect = Application("JOM_ConnectionString")
                    cn.Open(strConnect)

                    ' Function to fetch a random daily entry
                    Function GetRandomDictionaryEntry()
                        Dim rs, sql, today
                        Set rs = Server.CreateObject("ADODB.Recordset")
                        today = Day(Now())

                        ' Query to select a row based on a daily-changing ID
                        sql = "SELECT TOP 1 Word, Pronunciation, Partofspeech, Definition, Examplesentence, Category, Show " & _
                              "FROM Dictionary WHERE ID = " & (today Mod 100) & ";"
                        rs.Open sql, cn
                        If Not rs.EOF Then
                            GetRandomDictionaryEntry = "<p><strong>Word:</strong> " & rs("Word") & "</p>" & _
                                                       "<p><strong>Pronunciation:</strong> " & rs("Pronunciation") & "</p>" & _
                                                       "<p><strong>Part of Speech:</strong> " & rs("Partofspeech") & "</p>" & _
                                                       "<p><strong>Definition:</strong> " & rs("Definition") & "</p>" & _
                                                       "<p><strong>Example Sentence:</strong> " & rs("Examplesentence") & "</p>" & _
                                                       "<p><strong>Category:</strong> " & rs("Category") & "</p>" & _
                                                       "<p><strong>Show:</strong> " & rs("Show") & "</p>"
                        Else
                            GetRandomDictionaryEntry = "<p>No entry available for today.</p>"
                        End If
                        rs.Close
                        Set rs = Nothing
                    End Function

                    ' Fetch and display the random Dictionary entry
                    Response.Write GetRandomDictionaryEntry()
                    %>
                </div>
            </section>

            <!-- Random BlankGuy Entry -->
            <section class="entry-section">
                <div class="entry-title">Today's Random BlankGuy Entry</div>
                <div class="entry-content">
                    <%
                    ' Function to fetch a random daily entry for BlankGuy
                    Function GetRandomBlankGuyEntry()
                        Dim rs, sql, today
                        Set rs = Server.CreateObject("ADODB.Recordset")
                        today = Day(Now())

                        ' Query to select a row based on a daily-changing ID
                        sql = "SELECT TOP 1 Quote, [Ep#], Time, AudioFile, Show " & _
                              "FROM BlankGuy WHERE ID = " & (today Mod 100) & ";"
                        rs.Open sql, cn
                        If Not rs.EOF Then
                            Dim audioElement
                            If IsNull(rs("AudioFile")) Or rs("AudioFile") = "" Then
                                audioElement = "<p>No audio available</p>"
                            Else
                                audioElement = "<audio controls>" & _
                                               "<source src='/Audiofiles/" & rs("AudioFile") & "' type='audio/mpeg'>" & _
                                               "Your browser does not support the audio element.</audio>"
                            End If

                            GetRandomBlankGuyEntry = "<p><strong>Quote:</strong> " & rs("Quote") & "</p>" & _
                                                     "<p><strong>Episode:</strong> " & rs("Ep#") & "</p>" & _
                                                     "<p><strong>Time:</strong> " & rs("Time") & "</p>" & _
                                                     audioElement & _
                                                     "<p><strong>Show:</strong> " & rs("Show") & "</p>"
                        Else
                            GetRandomBlankGuyEntry = "<p>No entry available for today.</p>"
                        End If
                        rs.Close
                        Set rs = Nothing
                    End Function

                    ' Fetch and display the random BlankGuy entry
                    Response.Write GetRandomBlankGuyEntry()
                    %>
                </div>
            </section>
        </div>
    </main>
</body>

</html>

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
