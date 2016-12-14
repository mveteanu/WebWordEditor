<%@ Language=VBScript %>
<!-- #include file="scripts/adovbs.inc" -->
<%
 Set cn = Server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
 
 Set myCmd = Server.CreateObject("ADODB.Command")
 Set myCmd.ActiveConnection = cn
 myCmd.CommandText = "GetDocsInfo"
 myCmd.CommandType = adCmdStoredProc
 Set rs = myCmd.Execute
 Response.Write "id|Nume document|Numar imagini"& vbCrLf
 do until rs.EOF
   with Response
    .Write rs.Fields("id_doc").value & "|"
    .Write rs.Fields("nume").value & "|"
    .Write NullToZero(rs.Fields("nrimg").value) & vbCrLf
   end with
   rs.MoveNext
 loop        
 rs.Close
 Set rs = nothing
 Set myCmd = Nothing
 
 cn.Close
 Set cn=nothing
 
Function NullToZero(ByVal n)
 If IsNull(n) Then NullToZero = 0 Else NullToZero = n
End Function
%>