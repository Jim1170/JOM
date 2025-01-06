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
    <div class="page-heading">Comment Leavers</div>

    <!-- Contact Form Section -->
    <div class="contact-form-container">
        <p class="warning-message">
            Warning: Comment leaver card holders only. You will be turning in your regulation listener card if you proceed.
        </p>
        <form method="POST" action="" class="contact-form">
            <table>
                <tr>
                    <td><label for="formName">Name:</label></td>
                    <td><input type="text" id="formName" name="formName" required></td>
                </tr>
                <tr>
                    <td><label for="formEmail">Email:</label></td>
                    <td><input type="email" id="formEmail" name="formEmail" required></td>
                </tr>
                <tr>
                    <td><label for="formPhone">Phone (optional):</label></td>
                    <td><input type="tel" id="formPhone" name="formPhone"></td>
                </tr>
                <tr>
                    <td><label for="formComment">Comment:</label></td>
                    <td><textarea id="formComment" name="formComment" rows="5" required></textarea></td>
                </tr>
            </table>
            <div class="form-buttons">
                <input type="submit" value="Submit">
                <input type="reset" value="Reset">
            </div>
        </form>
    </div>

<%
' RETRIEVE DATA FROM FORM
formName = Request("formName")
formEmail = Request("formEmail")
formComment = Request("formComment")
formPhone = Request("formPhone")

' GET IP ADDRESS
Dim UserIPAddress
UserIPAddress = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If UserIPAddress = "" Then
    UserIPAddress = Request.ServerVariables("REMOTE_ADDR")
End If

' VALIDATE COMMENT
CommentOK = "yes"
If InStr(formComment, "http") > 0 Then
    CommentOK = "no"
End If

' IF DATA IS AVAILABLE, SEND EMAIL
If NOT (IsEmpty(formEmail)) AND CommentOK = "yes" Then
    sTo = "James@regulationlore.com.au"
    sFrom = "mark@tmtech.com.au"
    sSubject = "Regulation"
    sBody = "RegulationLore Contact Form has sent the following: Name: " & formName & vbCrLf & _
            "Email: " & formEmail & vbCrLf & _
            "Comment: " & formComment & vbCrLf & _
            "Phone: " & formPhone & vbCrLf & _
            "IP Address: " & UserIPAddress

    sBCC = "mark@tmtech.com.au"

    ' EMAIL CONFIGURATION AND SENDING
    Sub SendEmail()
        Set objMail = Server.CreateObject("CDO.Message")
        Set objConf = Server.CreateObject("CDO.Configuration")
        Set objFields = objConf.Fields

        With objFields
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.smtp2go.com"
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 2525
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "mark@tmtech.com.au"
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "mark2468"
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
            .Update
        End With

        With objMail
            Set .Configuration = objConf
            .From = sFrom
            .To = sTo
            .BCC = sBCC
            .Subject = sSubject
            .TextBody = sBody
        End With

        objMail.Send

        Set objFields = Nothing
        Set objConf = Nothing
        Set objMail = Nothing
    End Sub

    SendEmail()

    If Err.Number = 0 Then
        Response.Write("<p style='text-align: center; color: green;'>Thank you for your message. We will get back to you soon.</p>")
    Else
        Response.Write("<p style='text-align: center; color: red;'>Sorry, there was an error sending your message. Please try again later.</p>")
    End If
End If
%>

<!-- Styles -->
<style>
    .main-container {
        max-width: 800px;
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

    .warning-message {
        text-align: center;
        font-size: 1.2em;
        color: red;
        font-weight: bold;
        margin-bottom: 20px;
    }

    .contact-form-container {
        text-align: center;
    }

    .contact-form {
        margin: 0 auto;
        display: inline-block;
        text-align: left;
        width: 100%;
    }

    .contact-form table {
        width: 100%;
        border-spacing: 0 10px;
    }

    .contact-form td {
        padding: 5px;
        vertical-align: top;
    }

    .contact-form label {
        font-weight: bold;
    }

    .contact-form input[type=text], .contact-form input[type=email], .contact-form input[type=tel], .contact-form textarea {
        width: 100%;
        padding: 10px;
        border: 1px solid #ddd;
        border-radius: 4px;
        font-size: 1em;
    }

    .contact-form textarea {
        resize: vertical;
    }

    .form-buttons {
        text-align: center;
        margin-top: 15px;
    }

    .form-buttons input[type=submit], .form-buttons input[type=reset] {
        padding: 10px 20px;
        font-size: 1em;
        border: none;
        border-radius: 4px;
        cursor: pointer;
    }

    .form-buttons input[type=submit] {
        background-color: #28a745;
        color: #fff;
    }

    .form-buttons input[type=reset] {
        background-color: #dc3545;
        color: #fff;
    }

    .form-buttons input[type=submit]:hover {
        background-color: #218838;
    }

    .form-buttons input[type=reset]:hover {
        background-color: #c82333;
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
