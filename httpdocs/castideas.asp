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
    <div class="page-heading">Regulation Ideas</div>

    <div class="content-layout">
        <!-- Sidebar -->
        <div class="sidebar">
            <label for="statusFilter" class="sidebar-label">Filter by Status:</label>
            <select id="statusFilter" class="sidebar-select" onchange="filterTable()">
                <option value="all">All</option>
                <option value="complete.png">Complete</option>
                <option value="incomplete.png">Incomplete</option>
                <option value="partiallycomplete.png">Partially Complete</option>
                <option value="inprogress.png">In Progress</option>
            </select>
        </div>

        <!-- Main Content -->
        <div class="main-content">
            <table class="data-table">
                <caption>List of Regulation Ideas</caption>
                <tr>
                    <th scope="col">Regulation Idea</th>
                    <th scope="col">Description</th>
                    <th scope="col">Status</th>
                </tr>
                <%
                ' CREATE DATABASE CONNECTION
                Set cn = Server.CreateObject("ADODB.Connection")
                strConnect = Application("JOM_ConnectionString") ' Ensure this connection string is properly configured
                cn.Open(strConnect)

                ' SQL QUERY TO FETCH DATA FROM CASTIDEAS TABLE
                Set rs = cn.Execute("SELECT Castideas, Description, Status FROM Castideas ORDER BY Castideas ASC")
                While Not rs.EOF
                %>
                <tr>
                    <td><%= rs("Castideas") %></td>
                    <td><%= rs("Description") %></td>
                    <td>
                        <img src="/images/<%= rs("Status") %>" alt="<%= rs("Status") %>" class="status-image" />
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

    <!-- Floating Buttons -->
    <div class="floating-buttons">
        <button id="back-to-top" class="action-button" onclick="scrollToTop()">Back to Top</button>
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

    .sidebar-label {
        display: block;
        margin-bottom: 10px;
        font-weight: bold;
    }

    .sidebar-select {
        width: 100%;
        padding: 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
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
        background-color: #000;
        color: #fff;
        font-weight: bold;
    }

    .data-table td img.status-image {
        max-width: 150px;
        max-height: 150px;
    }

    .floating-buttons {
        position: fixed;
        bottom: 20px;
        left: 20px;
        display: none;
        flex-direction: column;
        gap: 10px;
    }

    .action-button {
        padding: 10px;
        background-color: #000;
        color: #fff;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        font-size: 1.1em;
    }

    .action-button:hover {
        background-color: #333;
    }
</style>

<!-- JavaScript -->
<script>
    function scrollToTop() {
        window.scrollTo({ top: 0, behavior: "smooth" });
    }

    function toggleFloatingButtons() {
        const floatingButtons = document.querySelector('.floating-buttons');
        if (window.scrollY > 200) {
            floatingButtons.style.display = 'flex';
        } else {
            floatingButtons.style.display = 'none';
        }
    }

    function filterTable() {
        const filterValue = document.getElementById("statusFilter").value.toLowerCase();
        const table = document.querySelector(".data-table");
        const rows = table.querySelectorAll("tr");

        rows.forEach((row, index) => {
            if (index === 0) return; // Skip the header row
            const statusCell = row.cells[2];
            const imgAltText = statusCell ? statusCell.querySelector("img").alt.toLowerCase() : "";

            if (filterValue === "all" || imgAltText === filterValue) {
                row.style.display = "";
            } else {
                row.style.display = "none";
            }
        });
    }

    window.addEventListener('scroll', toggleFloatingButtons);
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
