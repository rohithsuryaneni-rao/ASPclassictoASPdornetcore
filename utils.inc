<%
' This is a simple .inc file with common functions and constants

' Function to format a date
Function FormatDate(dt)
    If Not IsDate(dt) Then
        FormatDate = ""
    Else
        FormatDate = FormatDateTime(dt, vbShortDate)
    End If
End Function

' Function to trim and sanitize input
Function SanitizeInput(inputValue)
    If IsNull(inputValue) Then
        SanitizeInput = ""
    Else
        SanitizeInput = Trim(inputValue)
    End If
End Function

' Constant for database connection string
Const DB_CONNECTION_STRING = "Provider=SQLOLEDB;Data Source=myServerAddress;Initial Catalog=myDataBase;User Id=myUsername;Password=myPassword;"
%>