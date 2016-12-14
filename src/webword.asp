<%@ Language=VBScript %>
<!-- #include file="scripts/adovbs.inc" -->
<!-- #include file="scripts/HTMLEditControl.asp" -->
<%
Response.Buffer = True
Response.Expires = -1

Dim cn

Dim ServDocText
Dim ServDocID

If Request.QueryString.Count = 0 then
 ServDocText        = "Introduceti aici textul documentului..."
 ServDocID          = ""
Else
 set cn = Server.CreateObject("ADODB.Connection") 
 cn.Open Application("DSN")
 ServDocText = GetDOCText(Request.QueryString("DocID"), cn)
 If ServDocText = "" then ServDocText = "<font color=red>Document inexistent !</font>"
 ServDocID = "value=" & Request.QueryString("DocID")
 cn.Close
 set cn = nothing
End If
%>
<html>
<head>
  <title>WebWord Editor</title>
  <link rel="stylesheet" type="text/css" href="scripts/ptn.css">
</head>

<body unselectable="on" style="behavior:url('scripts/application.htc');">

<div id="WaitforForm" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Asteptati putin...
</td></tr></table>
</div>

<div id="Form1" style="overflow:hidden;visibility:hidden;" unselectable="on"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">

<div class=TForm style="width:265px;height:31px;" 
     style="left:6px;top:3px;"
     style="text-align:center;font-weight:bold;font-size:20px;color:#000064;">
VMA WebWord Editor</div>

<%
CreateChosePictureForm "scripts/docupload.asp", "target='FormReturn'", _
                       "<input type=hidden id='idopenpb' name='idopenpb' "& ServDocID &">" &_
                       "<input type=hidden id='delimglist' name='delimglist'>" &_
                       "<input type=hidden id='uploadlist' name='uploadlist'>" &_
                       "<input type=hidden id='doctext' name='doctext'>"

OpenHTMLEditControl 6,70,744,400, "scripts/HTMLEditor.htc", "images", "MyHTMLEdit"
OpenPage 705,375,""
Response.Write ServDocText
ClosePage
CloseHTMLEditControl
CreateHTMLPicturesToolBar 270,3
CreateHTMLZoomToolBar 609, 3
CreateHTMLEditToolBar 6,33
%>

<input id="ButtonSave" type=button value="Salveaza" title="Salveaza documentul"
       class=TButton style="width:75px;height:25px;"
       style="left:297px;top:485px;">
<input id="ButtonClose" type=button value="Inchide" title="Inchide fereastra"
       class=TButton style="width:75px;height:25px;"
       style="left:385px;top:485px;">
</div>

<div id="Form1Hidden" style="display:none;">
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language="vbscript" src="scripts/HTMLEditorUtils.vbs"></script>

<script language="vbscript">
Dim ServerImagesAtLoadTime, ServerTextAtLoadTime


Sub Window_onload
 Form1.style.visibility = "visible"
 WaitforForm.style.visibility = "hidden"

 ServerImagesAtLoadTime = GetEditPageServerImages(MyHTMLEdit_TextBox1)
 ServerTextAtLoadTime   = MyHTMLEdit_TextBox1.innerHTML
 MyHTMLEdit_TextBox1.ContentEditable = true
End Sub

' =====================================================

' Inchide documentul neconditionat
Sub CloseDocNeconditionat
  Window.returnValue = 0
  Window.close 
End Sub

' Inchide documentul salvat
Sub CloseSavedDoc
  Window.returnValue = 1
  window.close 
End Sub

' Evenimentul apare la apasarea butonului Close
Sub ButtonClose_onclick
  Dim wasChanged
  
  If ServerTextAtLoadTime = MyHTMLEdit_TextBox1.innerHTML then
    wasChanged = false
  Else
    wasChanged = true
  End If    

  If wasChanged then 
   Select case msgbox("Documentul s-a modificat. Doriti sa-l salvati?",vbYesNoCancel + vbInformation, "Atentie !")
     case vbYes ButtonSave_onclick
     case vbNo  CloseDocNeconditionat
   End Select
  Else
   CloseDocNeconditionat 
  End If
End Sub


' Evenimentul apare la apasarea butonului Save
Sub ButtonSave_onclick
  Dim deletimg, localimg
  Dim deletstring, localstring

  deletimg = ArrayDif(ServerImagesAtLoadTime, GetEditPageServerImages(MyHTMLEdit_TextBox1))
  localimg = GetEditPageLocalImages(MyHTMLEdit_TextBox1)

  If Join(deletimg)="" then
   deletstring = ""
  else
   deletstring = FileNamesArrayToIDCSV(deletimg, "FileID")
  end if

  CleanFilesForm ChosePictureFormular,localimg
  localstring = GetUploadFields(ChosePictureFormular)

  ChosePictureFormular.delimglist.value = deletstring
  ChosePictureFormular.uploadlist.value = localstring
  ChosePictureFormular.doctext.value = MyHTMLEdit_TextBox1.innerHTML
  ChosePictureFormular.Submit
End Sub
</script>

</body>
</html>

