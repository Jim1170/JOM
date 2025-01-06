<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Regulation Lore</title>

    <style>
        body {
            font-family: Arial, Helvetica, sans-serif;
        }

        /* Global styles */
        .container {
            width: 100%;
        }

        .header {
            text-align: center;
            background-color: #000;
            color: #fff;
            padding: 20px;
        }

        .nav-bar {
            display: flex;
            justify-content: space-around;
            background-color: #000;
            padding: 10px;
        }

        .nav-bar a {
            color: #fff;
            text-decoration: none;
            font-weight: bold;
        }

        .search-box {
            text-align: center;
            margin-top: 20px;
        }

        .word-container {
            margin: 10px 0;
        }

        .word {
            font-weight: bold;
            cursor: pointer;
            color: #000;
            margin: 5px 0;
        }

        .details {
            display: none;
            margin-left: 20px;
            background-color: #f0f0f0;
            padding: 10px;
            border-radius: 5px;
        }

        /* Footer */
        .footer {
            text-align: center;
            padding: 20px;
            font-size: 12px;
        }
    </style>

    <script>
        function toggleDetails(elementId) {
            const details = document.getElementById(elementId);
            details.style.display = details.style.display === "block" ? "none" : "block";
        }

        // Close details if clicking outside
        document.addEventListener('click', function (event) {
            const detailsElements = document.getElementsByClassName('details');
            let isClickInside = false;

            for (let i = 0; i < detailsElements.length; i++) {
                if (detailsElements[i].contains(event.target)) {
                    isClickInside = true;
                    break;
                }
            }

            if (!isClickInside && !event.target.classList.contains('word')) {
                for (let i = 0; i < detailsElements.length; i++) {
                    detailsElements[i].style.display = 'none';
                }
            }
        });
    </script>
</head>

<body>

    <div class="container">
        <!-- Header -->
        <div class="header">
            <a href="index.html">
                <h1>Regulation Lore</h1>
            </a>
        </div>

        <!-- Navigation -->
        <div class="nav-bar">
            <a href="dictionary.html">Dictionary</a>
            <a href="blankguy.html">I'm a ____ guy</a>
            <a href="regulationideas.html">Regulation Ideas</a>
            <a href="todo.html">To Do</a>
            <a href="collectables.html">Regulation Collectables</a>
        </div>

        <!-- Search Form -->
        <div class="search-box">
            <form method="post" action="">
                <label for="searchText">Enter Search Text and Click Search</label>
                <input id="searchText" name="SearchText" type="text" />
                <input name="Submit1" type="submit" value="Submit" />
            </form>
        </div>

        <!-- Word Results -->
        <div class="word-results">
            <table style="width: 30%" class="auto-style5">
                <thead>
                    <tr>
                        <th>Word</th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        ' GET VARIABLES FROM FORM
                        SearchText = Request("SearchText")
                        RecordCount = 0

                        ' CREATE DATABASE CONNECTION
                        Set cn = Server.CreateObject("ADODB.Connection")
                        strConnect = Application("JOM_ConnectionString")
                        cn.Open(strConnect)

                        ' SQL QUERY
                        Set rs = cn.Execute("SELECT * FROM Dictionary WHERE Word LIKE '%" & SearchText & "%'")

                        ' LOOP THROUGH RESULTS
                        Do While Not rs.EOF
                            uniqueId = "details-" & rs("ID")
                    %>

                    <tr>
                        <td>
                            <div class="word" onclick="toggleDetails('<%= uniqueId %>')">
                                <% Response.Write rs("word") %>
                            </div>
                            <div class="details" id="<%= uniqueId %>">
                                <p><strong>Pronunciation:</strong> <% Response.Write rs("Pronunciation") %></p>
                                <p><strong>Part of Speech:</strong> <% Response.Write rs("Partofspeech") %></p>
                                <p><strong>Definition:</strong> <% Response.Write rs("Definition") %></p>
                                <p><strong>Example Sentence:</strong> <% Response.Write rs("Examplesentence") %></p>
                            </div>
                        </td>
                    </tr>

                    <%
                            RecordCount = RecordCount + 1
                            rs.MoveNext
                        Loop
                        rs.Close
                        cn.Close
                    %>
                </tbody>
            </table>
        </div>

        <!-- Footer -->
        <div class="footer">
            <p>This is a fan-made site.</p>
            <a href="contactus.asp">Contact Us</a> | <a href="https://regulationstore.com/">Official Regulation Store</a>
        </div>
    </div>

</body>

</html>