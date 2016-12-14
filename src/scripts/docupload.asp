<%@ Language=VBScript %>
<!-- #include file="adovbs.inc" -->
<!-- #include file="FormUpload.asp" -->
<%
Dim FormFieldsArray
Dim StrMessageReturn

On Error Resume Next

FormFieldsArr  = GetFormItems
Set openpbid   = FormItem(FormFieldsArr, "idopenpb")
Set delimglist = FormItem(FormFieldsArr, "delimglist")
Set uploadlist = FormItem(FormFieldsArr, "uploadlist")
Set doctext    = FormItem(FormFieldsArr, "doctext")

Set cn = Server.CreateObject("ADODB.Connection")
cn.Open Application("DSN")
DeleteServerImages delimglist.value, cn
id_doc = EnterDOC(openpbid.value, cn)
imgenteredlist = EnterImages(uploadlist.value, id_doc, cn)
UpdateDocText ReplaceLocalImagesNames(doctext.value, "scripts/getattach_img.asp?FileID=", imgenteredlist), id_doc, cn
cn.Close 
Set cn = Nothing

If Err.number <> 0 then
  Err.Clear 
  StrMessageReturn = """Eroare la salvarea documentului in baza de date"", vbOkOnly+vbCritical"
Else
  StrMessageReturn = """Documentul a fost salvat cu succes"", vbOkOnly+vbInformation"
End If


' Sterge imaginile ce au ID-urile in lista CSV: servimglist
Sub DeleteServerImages(servimglist, objCon)
 Const SQLDel = "DELETE FROM TBAttachements WHERE id_attachement IN (@1)"
 If servimglist<>"" then objCon.Execute Replace(SQLDel, "@1", servimglist)
End Sub

' Introduce informatiile generale despre document  si returneaza
' id-ul documentului care tocmai s-a introdus
Function EnterDOC(idopenpb, objCon)
	Dim idpb
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "TBDocs", objCon, adOpenDynamic, adLockOptimistic, adCmdTable
	If idopenpb<>"" then 
	  rs.Filter = "id_doc = " & CStr(idopenpb)
	  idpb = rs.Fields("id_doc").Value 
	Else 
	  rs.AddNew
	  idpb = rs.Fields("id_doc").Value 
	  rs.Fields("nume").Value  = "Documentul " & CStr(idpb)
	End If  
	rs.Update 
	rs.Close 
	set rs   = nothing
	EnterDOC = idpb
End Function

' Se introduc imaginile in BD
Function EnterImages(imgupldlist, iddoc, objCon)
 Dim re
 
 re = ""
 If imgupldlist<>"" then
  Set rsa = Server.CreateObject("ADODB.Recordset")
  rsa.Open "TBAttachements", objCon, adOpenDynamic, adLockOptimistic, adCmdTable
  uploadfieldsarray = Split(imgupldlist,",",-1,1)
  for each uf in uploadfieldsarray
    set oFile = FormItem(FormFieldsArr, uf)
    rsa.AddNew
    rsa.Fields("id_doc").Value = iddoc
    rsa.Fields("attachementtype").Value = oFile.ContentType
    rsa.Fields("attachement").AppendChunk = oFile.BinaryData & ChrB(0)
    re = re & oFile.FileName & "|" & rsa.Fields("id_attachement").Value & Chr(1)
    rsa.Update
    set oFile = nothing
  next  
  rsa.Close
  set rsa = nothing
  re = Left(re, Len(re)-Len(Chr(1)))
 End If  
 
 EnterImages = re
End Function

' Inlocuieste numele imaginilor locale cu nume de imagini din BD
Function ReplaceLocalImagesNames(textpb, asppagename, limglist)
  Dim re
  re = textpb
  
  If limglist <> "" then
	ilst = Split(limglist, Chr(1), -1, 1)
	for each ii in ilst
	 ilstitem = Split(ii, "|", -1, 1)
	 re = Replace(re, ilstitem(0), asppagename & ilstitem(1))
	next
  End If

  ReplaceLocalImagesNames = re
End Function

' Se updateaza un document proaspat introdus in vederea
' adaugarii si a textului
Sub UpdateDocText(textpb, iddoc, objCon)
 Const SQLSel = "SELECT * FROM TBDocs WHERE id_doc=@1"
 set rsu = Server.CreateObject("ADODB.Recordset")
 rsu.Open Replace(SQLSel,"@1",CStr(iddoc)), objCon, adOpenDynamic, adLockOptimistic, adCmdText
 rsu.Fields("text").Value = textpb
 rsu.Update 
 rsu.Close 
 set rsu = nothing
End Sub
%>
<body>
<script language=vbscript>
sub window_onload
 msgbox <%=StrMessageReturn%>
 call window.parent.CloseSavedDoc
end sub
</script>
</body>