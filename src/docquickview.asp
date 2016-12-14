<%@ Language=VBScript %>
<!-- #include file="scripts/adovbs.inc" -->
<%
 Response.Buffer = True
 Response.Expires = -1
 
 Dim DocsHTML, cn
 
 set cn = Server.CreateObject("ADODB.Connection") 
 cn.Open Application("DSN")
 DocsHTML = GetCompleteDocsHTML(Request.QueryString("DocIDs"), cn)
 cn.Close
 set cn = nothing

 ' Intoarce textul HTML COMPLET al unei pagini cu o problema
 Function GetCompleteDocsHTML(DocIDs, objCon)
  Const SQLSel = "SELECT * FROM TBDocs WHERE id_doc IN (@1)"
  Dim DocDivs, contor
  Dim rs, re

  DocDivs = "<div class='THTMLEditPageBorder' style='width:645px;'>" & vbCrLf &_
            "<div class='THTMLEditPage' style='width:645px;'>" & vbCrLf &_
            "<div style='display:block;padding:5px;'><b>@1. @2</b></div>" & vbCrLf &_
            "<div style='display:block;padding:5px;'>@3</div><br>" & vbCrLf &_
            "</div>" & vbCrLf &_
            "</div>" & vbCrLf

  set rs = objCon.Execute(Replace(SQLSel, "@1", DocIDs))
  contor = 1: re = ""
  do until rs.EOF
    re = re & Replace(Replace(Replace(DocDivs, "@3", rs.Fields("text").value), "@2", rs.Fields("nume").value), "@1", CStr(contor))
    contor = contor + 1
    rs.movenext
  loop
  rs.Close
  set rs = nothing
  
  GetCompleteDocsHTML = re
 End Function
%>
<html>
<head>
  <title>Vizualizare documente</title>
  <link rel="stylesheet" type="text/css" href="scripts/ptn.css">
</head>
<body>

<div id="WaitforForm" style="overflow:hidden;visibility:visible;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">
<table border=0 width=100% height=100%><tr><td align=center valign=center>
Asteptati putin...
</td></tr></table>
</div>

<div id="Form1" style="overflow:hidden;visibility:hidden;"
     class="TForm" style="border: none;"
     style="left:0px;top:0px;width:100%;height:100%;">

<fieldset id=GroupBox1 class=TGroupBox 
          style="width:698px;height:416px;"
          style="left:7px;top:8px;">
<legend>Documente</legend>
<div id="ViewPbDIV" style="position:absolute; left:8px; top:16px; width:680px; height:392px; background-color:threedshadow; border:inset thin; FONT-FAMILY:Times New Roman; FONT-Size: 12pt; overflow:auto;">
<%=DocsHTML%>
</div>
</fieldset>

<input id="Button1" type=button value="Inchide" title="Inchide fereastra"
       class=TButton style="width:75px;height:25px;"
       style="left:311px;top:440px;">
</div>

<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button1_onclick
  Window.close 
End Sub
</script>

</body>
</html>