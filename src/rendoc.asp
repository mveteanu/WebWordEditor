<%@ Language=VBScript %>
<!-- #include file="scripts/adovbs.inc" -->
<%
Response.Buffer = True
Response.Expires = -1

set cn = Server.CreateObject("ADODB.Connection") 
cn.Open Application("DSN")
PrintRenameForm(GetDocumentName(Request.QueryString("DocID"), cn))
cn.Close 
set cn = nothing

' Intoarce numele documentului
Function GetDocumentName(DocID, objCon)
 Const SQLSel = "SELECT * FROM TBDocs WHERE id_doc=@1"
 Dim re, Rs
 Set Rs = objCon.Execute(Replace(SQLSel, "@1", DocID))
 If not Rs.EOF then re = Rs.Fields("nume").Value Else re = ""  
 Rs.Close
 set rs = nothing
 GetDocumentName = re
End Function


Sub PrintRenameForm(numevechi)
%>
<html>
<head>
  <title>Redenumire document</title>
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
<span class=TLabel style="width:31px;height:13px;"
      style="left:8px;top:16px;">
Nume:
</span>
<input id="Edit1" type=text maxlength=50 value="<%=numevechi%>"
       class=TEdit style="width:257px;height:21px;"
       style="left:8px;top:32px;">
<input id="Button1" type=button value="Accepta" title="Accepta modificarile"
       class=TButton style="width:75px;height:25px;"
       style="left:55px;top:80px;">
<input id="Button2" type=button value="Renunta" title="Anuleaza modificarile facute"
       class=TButton style="width:75px;height:25px;"
       style="left:143px;top:80px;">
</div>


<script language=vbscript>
' Evenimentul apare la incarcarea documentului
Sub window_onload
  Form1.style.visibility = "visible"
  WaitforForm.style.visibility = "hidden"
  Edit1.focus 
End Sub

' Trateaza evenimentul aparut la apasarea butonului OK
Sub Button1_onclick
  If Trim(Edit1.value) = "" then 
    MsgBox "Numele documentului nu poate fi nul sau spatiu.", vbOkOnly+vbExclamation, "Atentie!"
  Else
    window.returnValue = Edit1.value
    window.close 
  End If  
End Sub

' Trateaza evenimentul aparut la apasarea butonului Cancel
Sub Button2_onclick
  Window.close 
End Sub
</script>

</body>
</html>
<%
End Sub
%>
