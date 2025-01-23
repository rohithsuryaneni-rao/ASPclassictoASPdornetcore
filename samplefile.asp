<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim userName
    userName = Request.Form("userName")
    Response.Write("<h2>Hello, " & Server.HTMLEncode(userName) & "!</h2>")
End If
%>

<!DOCTYPE html>
<html>
<head>
    <title>ASP Classic Example</title>
</head>
<body>
    <h1>Welcome to ASP Classic</h1>
    <form method="post" action="example.asp">
        <label for="userName">Enter your name:</label>
        <input type="text" id="userName" name="userName" required>
        <button type="submit">Submit</button>
    </form>
</body>
</html>
