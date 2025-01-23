<%
' Retrieve the file name from the query string
Dim fileName
fileName = Request.QueryString("file")

' Define the path to the files directory
Dim filePath
filePath = Server.MapPath("/files/" & fileName)

' Create the FileSystemObject
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

' Check if the file exists
If objFSO.FileExists(filePath) Then
    ' Open the file as a binary stream
    Dim objStream
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1 ' adTypeBinary
    objStream.LoadFromFile(filePath)
    
    ' Set the appropriate headers to prompt the download
    Response.Clear
    Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
    Response.AddHeader "Content-Length", objStream.Size
    Response.ContentType = "application/octet-stream"
    
    ' Send the binary data to the client
    Response.BinaryWrite objStream.Read
    
    ' Clean up
    objStream.Close
    Set objStream = Nothing
Else
    ' Handle the case where the file does not exist
    Response.Write "File not found."
End If

' Clean up
Set objFSO = Nothing
%>
