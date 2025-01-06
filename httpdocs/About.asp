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
    <div class="page-heading">About</div>

    <!-- Welcome Section -->
    <div class="welcome-section">
        <p>
            Welcome to <strong>regulationlore.com.au!</strong> A fan-made site bringing you the deep lore behind the nothing.
        </p>
        <p>
            Have you ever asked yourself, <em>What's a bean hole?</em> or <em>Who is Leona Waddell?</em> You can find the answers to these, and many other questions that no one has ever asked, right here on
            <a href="http://regulationlore.com.au/" target="_blank" class="external-link">regulationlore.com.au</a>.
        </p>
        <p>
            This site was made as a project to learn website building and design. It is still under construction and will continue to grow. This is a fan-made site aimed at providing context for the podcast content produced by the Regulation Podcast.
        </p>
        <p>
            If you wish to get in touch or leave a comment, visit the
            <a href="contactus.asp" class="internal-link">comment leavers</a> page.
        </p>
    </div>

    <!-- Links to Official Content -->
    <div class="links-section">
        <p>For official content or community links, see below:</p>
        <ul class="official-links">
            <li><a href="https://podcasts.apple.com/au/podcast/regulation-podcast/id1515027050" target="_blank" class="external-link">Regulation Podcast</a></li>
            <li><a href="https://www.youtube.com/channel/UCPP6JksBPYY3MeFQn5yanIw" target="_blank" class="external-link">Regulation YouTube</a></li>
            <li><a href="https://www.youtube.com/@TheRegulationPodcast" target="_blank" class="external-link">Regulation Gameplay YouTube</a></li>
            <li><a href="https://regulationstore.com/" target="_blank" class="external-link">Regulation Store</a></li>
            <li><a href="https://www.patreon.com/TheRegulationPod" target="_blank" class="external-link">Regulatreon - Regulation Patreon</a></li>
            <li><a href="https://www.twitch.tv/theregulationpod" target="_blank" class="external-link">Regulation Twitch</a></li>
            <li><a href="https://www.instagram.com/fuckfacepod/?hl=en" target="_blank" class="external-link">Regulation Instagram</a></li>
            <li><a href="https://www.reddit.com/r/theregulationpod/?rdt=39997" target="_blank" class="external-link">Regulation Reddit</a></li>
            <li><a href="https://www.patreon.com/posts/join-our-discord-104064126" target="_blank" class="external-link">Regulation Discord</a></li>
        </ul>
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

    .welcome-section {
        text-align: center;
        margin: 20px auto;
        font-size: 1.2em;
        line-height: 1.6;
        color: #333;
    }

    .links-section {
        text-align: center;
        margin: 20px auto;
        font-size: 1em;
    }

    .official-links {
        list-style: none;
        padding: 0;
        margin: 20px 0;
    }

    .official-links li {
        margin-bottom: 10px;
    }

    .external-link, .internal-link {
        color: #0066cc;
        text-decoration: none;
    }

    .external-link:hover, .internal-link:hover {
        text-decoration: underline;
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
