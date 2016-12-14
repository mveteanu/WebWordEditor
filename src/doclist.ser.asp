<%@ Language=VBScript %>
<!-- #include file="scripts/adovbs.inc" -->
<%
 Const MsgDelOK = """Documentele specificate au fost sterse cu succes"", vbOkOnly+vbInformation"
 Const MsgRenOK = """Documentul specificat a fost redenumit cu succes"", vbOkOnly+vbInformation"
 Const MsgError = """Eroare la operatiile cu BD"", vbOkOnly+vbCritical"
 Dim Mesaj, ReqAct
 
 On Error Resume Next
 
 set cn = server.CreateObject("ADODB.Connection")
 cn.Open Application("DSN")
  
 ReqAct = Request.Form("SelectAction")
 if ReqAct="del" then
   DeleteDocuments Request.Form("SelectList"), cn
   Mesaj = MsgDelOK
 elseif ReqAct="ren" then
   RenameDocument Request.Form("SelectList"), Request.Form("SelectValue"), cn
   Mesaj = MsgRenOK
 end if
 
 cn.Close
 set cn = nothing
 
 if Err.number <> 0 then
   Mesaj = MsgError
   Err.Clear 
 end if
 
 Sub DeleteDocuments(docsids, objCon)
   Const SQLDel = "DELETE FROM TBDocs WHERE id_doc IN (@1)"
   objCon.Execute Replace(SQLDel, "@1", docsids)
 End Sub

 Sub RenameDocument(docid, newname, objCon)
   Const SQLRen = "UPDATE TBDocs SET nume = '@1' WHERE id_doc = @2"
   objCon.Execute Replace(Replace(SQLRen, "@2", CStr(docid)), "@1", newname)
 End Sub
%>
<body>
<script language=vbscript>
  window.parent.ReloadTDC
  msgbox <%=Mesaj%>
</script>
</body>
