<%
' Get the search text from the user
SearchText = Request("SearchText")

' Create the database connection
Set cn = Server.CreateObject("ADODB.Connection")
strConnect = Application("JOM_ConnectionString")
cn.Open(strConnect)

' Prepare the SQL query with placeholders for the user input
sqlQuery = "SELECT * FROM Dictionary WHERE Word LIKE ? OR Pronunciation LIKE ? OR Partofspeech LIKE ? OR Definition LIKE ? OR Examplesentence LIKE ?"

' Create the command object
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = cn
cmd.CommandText = sqlQuery

' Add parameters to the command
cmd.Parameters.Append cmd.CreateParameter(, 8, 1, 255, "%" & SearchText & "%")  ' 8 = adVarChar, 1 = adParamInput
cmd.Parameters.Append cmd.CreateParameter(, 8, 1, 255, "%" & SearchText & "%")
cmd.Parameters.Append cmd.CreateParameter(, 8, 1, 255, "%" & SearchText & "%")
cmd.Parameters.Append cmd.CreateParameter(, 8, 1, 255, "%" & SearchText & "%")
cmd.Parameters.Append cmd.CreateParameter(, 8, 1, 255, "%" & SearchText & "%")

' Execute the query
Set rs = cmd.Execute()

' Loop through the results
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
    rs.MoveNext
Loop
rs.Close
cn.Close
%>
