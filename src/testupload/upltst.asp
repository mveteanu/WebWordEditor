<%@ Language=VBScript %>
<%
If Request.TotalBytes=0 Then Call PrintUploadForm Else Call PrintUploadedData

Sub PrintUploadedData
  Response.Write "<head><title>Test FormUpload</title></head><body><pre>" & vbCrLf
  Response.BinaryWrite(Request.BinaryRead(Request.TotalBytes))
  Response.Write "</pre></body>"
End Sub

Sub PrintUploadForm
%>
<head><title>Test FormUpload</title></head><body>
<form method=post enctype="multipart/form-data" action="upltst.asp">
Nume:<br><input type=text name="nume"><br><br>
Comentarii:<br><textarea cols=40 rows=5 name="comentarii"></textarea><br><br>
Imagine:<br><input type=file name="imagine"><br><br>
<input type=submit name="submit" value="Trimite">
</form></body>
<%
End Sub
%>
