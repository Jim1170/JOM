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
    <div class="page-heading">To Do Lists</div>

    <!-- To-Do Lists Section -->
    <div class="content-layout">
        <!-- Vancouver To-Do List -->
        <div class="list-wrapper">
            <table class="data-table">
                <caption>Vancouver To-Do</caption>
                <tr>
                    <th scope="col" class="checkbox-cell">Done</th>
                    <th scope="col">Task</th>
                </tr>
                <%
                ' CREATE DATABASE CONNECTION
                Set cn = Server.CreateObject("ADODB.Connection")
                strConnect = Application("JOM_ConnectionString") ' Ensure this connection string is properly configured
                cn.Open(strConnect)

                ' SQL QUERY TO FETCH DATA FROM THE Vancouvertodo TABLE
                Set rs = cn.Execute("SELECT VancouverToDo FROM Vancouvertodo")
                While Not rs.EOF
                %>
                <tr>
                    <td class="checkbox-cell">
                        <input type="checkbox" onclick="toggleStrikeThrough(this)" />
                    </td>
                    <td>
                        <span><%= rs("VancouverToDo") %></span>
                    </td>
                </tr>
                <%
                    rs.MoveNext
                Wend
                rs.Close
                %>
            </table>
        </div>

        <!-- Andrew's Big Day List -->
        <div class="list-wrapper">
            <table class="data-table">
                <caption>Andrew's Big Day</caption>
                <tr>
                    <th scope="col" class="checkbox-cell">Done</th>
                    <th scope="col">Task</th>
                </tr>
                <%
                ' SQL QUERY TO FETCH DATA FROM THE Andrewsbigday TABLE
                Set rs = cn.Execute("SELECT AndrewsBigDay FROM Andrewsbigday")
                While Not rs.EOF
                %>
                <tr>
                    <td class="checkbox-cell">
                        <input type="checkbox" onclick="toggleStrikeThrough(this)" />
                    </td>
                    <td>
                        <span><%= rs("AndrewsBigDay") %></span>
                    </td>
                </tr>
                <%
                    rs.MoveNext
                Wend
                rs.Close
                cn.Close
                %>
            </table>
        </div>
    </div>
</div>

<!-- JavaScript for Interactive Checkboxes -->
<script type="text/javascript">
    function toggleStrikeThrough(checkbox) {
        const textElement = checkbox.parentElement.nextElementSibling.querySelector("span");
        if (checkbox.checked) {
            textElement.classList.add("completed");
        } else {
            textElement.classList.remove("completed");
        }
    }
</script>

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
        justify-content: space-between;
    }

    .list-wrapper {
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
        text-align: left;
    }

    .data-table th {
        background-color: #000;
        color: #fff;
        font-weight: bold;
        text-align: center;
    }

    .checkbox-cell {
        width: 30px;
        text-align: center;
    }

    .completed {
        text-decoration: line-through;
        color: gray;
    }
</style>

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
