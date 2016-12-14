<%@ Language=VBScript %>
<!-- #include file="scripts/TableControl.asp" -->
<%Application("DSN") = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("db\webword.mdb")%>

<HTML>
<head>
 <title>Test WebWord Editor</title>
 <link rel="stylesheet" type="text/css" href="scripts/ptn.css">
</head>
<BODY unselectable="on" style="behavior:url('scripts/application.htc');">

<script language="vbscript" src="scripts/menu.vbs"></script>
<script language="vbscript" src="scripts/TableControlEvents.vbs"></script>

<div id="Form1" unselectable='on'
     class=TForm style="width:602px;height:254px;"
     style="left:Expression((document.body.clientWidth/2)-(this.offsetWidth/2));top:80px;">
<%CreateTableControl 8, 8, 201, Array("Nume document", "Numar imagini"), Array(390,190), 2, "doclist.dat.asp" , true, "MyTable"%>
<button id="Button1" title="Selectare documente"
       class=TButton style="width:85px;height:25px;"
       style="left:8px;top:216px;">Selectare <font face="Webdings">6</font></button>
<input id="Button2" type=button value="Redenumeste" title="Redenumeste un document"
       class=TButton style="width:85px;height:25px;"
       style="left:106px;top:216px;">
<input id="Button3" type=button value="Sterge" title="Sterge un document"
       class=TButton style="width:85px;height:25px;"
       style="left:205px;top:216px;">
<input id="Button4" type=button value="Adauga doc." title="Creaza un nou document"
       class=TButton style="width:85px;height:25px;"
       style="left:303px;top:216px;">
<input id="Button5" type=button value="Vizualizare" title="Vizualizeaza rapid mai multe documente"
       class=TButton style="width:85px;height:25px;"
       style="left:402px;top:216px;">
<input id="Button6" type=button value="Deschide doc." title="Deschide un document pentru editare"
       class=TButton style="width:85px;height:25px;"
       style="left:500px;top:216px;">
</div>

<div id="Form1Hidden" style="display:none;">
<form name="FormularH" method="post" action="doclist.ser.asp" target="FormReturn">
<input type=text id="SelectAction" name="SelectAction">
<input type=text id="SelectList" name="SelectList">
<input type=text id="SelectValue" name="SelectValue">
</form>
<IFRAME ID=FormReturn Name=FormReturn FRAMEBORDER=No FRAMESPACING=0 width=100% scrolling=no>
</IFRAME>
</div>

<script language=vbscript>
Dim mymenuselect
MenuItems1 = Array("Selecteaza pe toate","Deselecteaza pe toate")

' Activeaza/dezactiveaza butoanele
Sub ActivateButtons(state)
 For i = 1 to 6
   Form1.all("Button" & CStr(i)).disabled = not state
 Next
End Sub

' Apare la incarcarea datelor in TDC
Sub tdcMyTable_ondatasetcomplete
  ActivateButtons true
End Sub

' Determina reincarcarea TDC-ului
Sub ReloadTDC
  tdcMyTable.DataURL = tdcMyTable.DataURL
  tdcMyTable.Reset
End Sub

' Converteste un string de tipul 27px intr-un intreg de tipul 27
Function StyleSizeToInt(ssize)
  StyleSizeToInt = CInt(Left(ssize,Len(ssize)-2))
End Function

' Intoarce in format CSV ID-urile recordurilor selectate si afiseaza un
' mesaj de avertizare daca nu s-a selectat nici o inregistrare
Function GetSelectedRecords
  Dim RecList
  RecList = TableGetSelected(tblMyTable)
  if RecList = "" then  MsgBox "Trebuie sa selectati cel putin o inregistrare inainte de a continua.", vbOkOnly+vbExclamation
  GetSelectedRecords = RecList
End Function

' Intoarce sub forma de string ID-ul recordului selectat.
' Daca nu se selecteaza nici o inregistrare sau se selecteaza mai mult de una
' atunci se afiseaza un mesaj si se intoarce sirul vid.
Function GetSelectedRecord
  Dim RecList
  Dim RecArray
  
  RecList  = TableGetSelected(tblMyTable)
  RecArray = Split(RecList,",",-1,1)
  If (UBound(RecArray)-LBound(RecArray))<>0 then 
    MsgBox "Trebuie sa selectati o singura inregistrare din lista.", vbOkOnly+vbExclamation
    RecList = ""
  End If  
  GetSelectedRecord = RecList
End Function

Sub handlemenuselectclick(html)
 if html="<HR>" or html="" then exit sub
 mymenuselect.Hide
 set mymenuselect=nothing

 select case html     
  case MenuItems1(0) TableSelectAll tblMyTable, true
  case MenuItems1(1) TableSelectAll tblMyTable, false
 end select
End Sub

' Evenimentul apare la apasarea butonului Selectare
Sub Button1_OnClick
 Dim leftm, topm
   
 leftm = 3 + StyleSizeToInt(Button1.style.left) + StyleSizeToInt(Form1.style.left)
 topm  = 3 + StyleSizeToInt(Button1.style.top)  + StyleSizeToInt(Button1.style.height) + StyleSizeToInt(Form1.style.top)
 set mymenuselect = showmenu(leftm, topm, 140, "handlemenuselectclick", MenuItems1)
End Sub

' Evenimentul apare la apasarea butonului Rename
Sub Button2_onclick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub
  
  DocNewName = ShowModalDialog("rendoc.asp?DocID=" & RecList, , "dialogWidth=280px;dialogHeight=150px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no")
  If DocNewName<>"" then
   FormularH.SelectAction.value = "ren"
   FormularH.SelectList.value = RecList
   FormularH.SelectValue.value = DocNewName
   ActivateButtons false
   FormularH.submit 
  End If 
End Sub

' Trateaza evenimentul OnClick la butonul Delete Documents
Sub Button3_onclick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  if msgbox("Sunteti sigur ca doriti sa stergeti documentele selectate?",vbYesNo+vbQuestion,"Confirmati") = vbNo then Exit Sub

  FormularH.SelectAction.value = "del"
  FormularH.SelectList.value = RecList
  ActivateButtons false
  FormularH.submit 
End Sub

' Trateaza evenimentul OnClick la butonul Add Document
Sub Button4_onclick
  r = CInt(ShowModalDialog("webword.asp", , "dialogWidth=768px;dialogHeight=547px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"))
  If r = 1 then ReloadTDC
End Sub

' Trateaza evenimentul OnClick la butonul Quick View Multiple Documents
Sub Button5_OnClick
  Dim RecList
  RecList=GetSelectedRecords
  if RecList="" then Exit Sub
  
  ShowModalDialog "docquickview.asp?DocIDs=" & RecList, "", "dialogWidth=720px;dialogHeight=500px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"
End Sub

' Trateaza evenimentul OnClick la butonul Edit Document
Sub Button6_OnClick
  Dim RecList
  RecList=GetSelectedRecord
  if RecList="" then Exit Sub

  r = CInt(ShowModalDialog("webword.asp?DocID=" & RecList, , "dialogWidth=768px;dialogHeight=547px; scrollbars=no; scroll=no; center=yes; border=thin; help=no; status=no"))
  If r = 1 then ReloadTDC
End Sub
</script>

</BODY>
</HTML>
