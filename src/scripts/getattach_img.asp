<%@ Language=VBScript %>
<!-- #include file="adovbs.inc" -->
<%
Const SQLSel = "SELECT * FROM TBAttachements WHERE id_attachement=@1"
Dim cn, rs, fisid

fisid = Request.QueryString("FileID")
If fisid<>"" then
	set cn = Server.CreateObject("ADODB.Connection")
	set rs = Server.CreateObject("ADODB.Recordset")
	cn.CursorLocation = adUseClient
	cn.Open Application("DSN")
	rs.ActiveConnection = cn
	rs.Open Replace(SQLSel, "@1", fisid)
    If not rs.EOF then
      Response.ContentType = rs.Fields("attachementtype")
      Response.BinaryWrite rs.Fields("attachement")
    else
      PrintEndError
    end if
	rs.Close
	cn.Close
	set cn = nothing
	set rs = nothing
Else
    PrintEndError 
End If    

Sub PrintEndError
    Response.Write "<b>Eroare:</b> Nu ati specificat FileID sau fisierul cu ID-ul dat nu se gaseste in baza de date."
    Response.End
End Sub
%>