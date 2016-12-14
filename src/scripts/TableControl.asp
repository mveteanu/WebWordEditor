<!-- #include file="ControlUtils.asp" -->
<%
' *********************************************************************
' Control grafic de tipul TTableGrid realizat prin generarea pe server
' in ASP si folosind script client VBScript pentru interactiune in Browser
' Autor: Marian Veteanu
' Data:  15 martie 2001
' *********************************************************************

' ***************************************************
' BEGIN
' PUBLIC
' SECTION
' ***************************************************

' Subrutina principala care construieste tabelul in pagina
' impreuna cu TDC-ul asociat (vezi info despre TDC mai jos)
' Intrare:
'   iLeft   \ iLeft si iTop sunt folosite pentru pozitionarea controlului
'   iTop    / daca acestea lipsesc atunci controlul este pus cu pozitie relativa
'   iHeight = inaltimea controlului (lungimea se  calculeaza ca suma a lungimilor coloanelor)
'   strTabNames = Array cu denumirile coloanelor (trebuie sa coincida cu cele din sursa de date pentru TDC !!!)
'   iTabSizes = Array cu marimile pe orizontala a coloanelor. Array-ul trebuie sa aiba acelasi numar de elemente ca precedentul !!!
'   iTipTable = Tipul tabelului (0=fara interactiune, 1=cu selectie simpla, 2=cu selectie multipla)
'   tdcURL = URL-ul sau fisierul de la care obiectul TDC primeste date
'   bAllowSorting = Daca este true atunci se permite sortarea datelor din tabel pe Client prin actionarea headerlor coloanelor
'   strComponentName = Numele componentei - este folosit pentru generarea ID-urilor celorlalte elemente implicate in construirea controlului
' Format TDC Data: 
'   Datele primite de TDC trebuie sa fie sub urmatoarea forma:
'     id|nume|prenume|email
'     1|Veteanu|Marian|mveteanu@yahoo.com
'     etc.
'   Campul ID este obligatoriu daca se doreste obtinerea liniilor din tabel care au fost selectate de utilizator (in cazul cand iTipTable=1 sau 2)
'   Numele campurilor datelor primite de TDC trebuie sa coincida cu cele specificate in Array-ul strTabNames
Public Sub CreateTableControl(iLeft, iTop, iHeight, strTabNames, iTabSizes, iTipTable, tdcURL, bAllowSorting, strComponentName)
  Dim i

  ControlName = strComponentName
  ControlLeft = iLeft
  ControlTop = iTop
  ControlWidth = 0
  ControlHeight = iHeight
  ControlTabNames = strTabNames
  ControlTabSizes = iTabSizes
  ControlAllowSorting = bAllowSorting
  
  for each i in iTabSizes 
    ControlWidth = ControlWidth + i
  next
  ControlWidth = ControlWidth
  
  call AddTDC("tdc" & ControlName, "|", tdcURL)
  
  call OpenCanvas("table" & strComponentName, ControlLeft, ControlTop, ControlWidth, ControlHeight)
  for i = 0 to UBound(ControlTabNames)
    call CreateHeaderTab(i)
  next
  
  call CreateTableBody(iTipTable)
  call CloseCanvas
End Sub


'
' BEGIN PRIVATE SECTION
'

Dim ControlName
Dim ControlLeft
Dim ControlTop
Dim ControlWidth
Dim ControlHeight
Dim ControlTabNames
Dim ControlTabSizes
Dim ControlAllowSorting


' Creaza o patratica pe headerul tabelului. Aceasta reprezinta
' headerul unei coloane
Private Sub CreateHeaderTab(iTab)
  Dim DIVTemplate
  Dim itabLeft
  Dim i
  
  DIVTemplate = "<DIV id='@1' @6 unselectable='on' style='POSITION: absolute; OVERFLOW: hidden; CURSOR: @7; BORDER: outset thin; LEFT:@2px; TOP:0px; WIDTH:@3px; HEIGHT:20px; BACKGROUND-COLOR: buttonface; FONT-FAMILY:verdana; FONT-SIZE:8pt; FONT-WEIGHT:bolder; PADDING:1px; TEXT-ALIGN:left;'>@5@4</DIV>"

  iTabLeft = 0
  for i=0 to iTab-1
    iTabLeft = iTabLeft + ControlTabSizes(i)
  next
  
  DIVTemplate = Replace(DIVTemplate,"@1","tab" & CvtStrToID(ControlTabNames(iTab)))
  DIVTemplate = Replace(DIVTemplate,"@2",CStr(iTabLeft))
  DIVTemplate = Replace(DIVTemplate,"@3",CStr(ControlTabSizes(iTab)))
  if ControlAllowSorting then
   DIVTemplate = Replace(DIVTemplate,"@4","<span unselectable='on' style='position:absolute;top:2px;'>"& ControlTabNames(iTab)& "</span>")
   DIVTemplate = Replace(DIVTemplate,"@5","<span unselectable='on' style='display:none;font-family:webdings;'>6</span>")
   DIVTemplate = Replace(DIVTemplate,"@6","onclick='HandleTableHeaderSort tdc"& ControlName &", "& iTab &", "& UBound(ControlTabNames) &"'")
   DIVTemplate = Replace(DIVTemplate,"@7","hand")
  else
   DIVTemplate = Replace(DIVTemplate,"@4",ControlTabNames(iTab))
   DIVTemplate = Replace(DIVTemplate,"@5","")
   DIVTemplate = Replace(DIVTemplate,"@6","")
   DIVTemplate = Replace(DIVTemplate,"@7","default")
  end if
  Response.Write DIVTemplate & vbCrLf
End Sub



' Creaza partea prinipala a tabelului (div-ul cu autoscroll si tabelul cu
' incarcare din TDC
Private Sub CreateTableBody(tipcontrol)
  Dim DIVTemplate1
  Dim DIVTemplate2
  Dim TableTemplate1
  Dim TableTemplate2
  Dim TableTemplate3
  Dim TipControlStr
  Dim i
  
  DIVTemplate1 = "<DIV unselectable='on' style='BORDER: outset thin; POSITION: absolute; LEFT:0px; TOP:20px; WIDTH:@1px; HEIGHT:@2px; OVERFLOW:auto;'>"
  DIVTemplate2 = "</DIV>"
  DIVTemplate1 = Replace(DIVTemplate1,"@1",ControlWidth)
  DIVTemplate1 = Replace(DIVTemplate1,"@2",ControlHeight-20)
  
  select case tipcontrol
    case 0 TipControlStr = "style='cursor:default;'"
    case 1 TipControlStr = "style='cursor:hand;' onclick='vbscript:HandleTableClick(1)'"
    case 2 TipControlStr = "style='cursor:hand;' onclick='vbscript:HandleTableClick(2)'"
  end select
  
  TableTemplate1 = "<TABLE "& TipControlStr &" id='tbl"& ControlName &"' datasrc=#tdc"& ControlName &" class='TTableGrid' border=0 width=100% bgcolor=white cellspacing=0 cellpadding=0>" &vbcrlf &_
                   "<TBODY>" &vbcrlf &_
                   "<TR>" &vbcrlf
 
  TableTemplate2 = "</TR>" &vbcrlf &_
                   "</TBODY>" &vbcrlf &_
                   "</TABLE>" &vbcrlf
  
  TableTemplate3 = "<TD align=left valign=center width="& CStr(ControlTabSizes(0)) &" height=20 class='TTableRowUnSelected'><span datafld='id' style='display:none;'></span><span unselectable='on' DATAFORMATAS=HTML style='overflow:hidden;width:"& CStr(ControlTabSizes(i)-4) &"px;' datafld='"& ControlTabNames(0) &"'></span></TD>" &vbcrlf
  for i = 1 to UBound(ControlTabNames)
    TableTemplate3 = TableTemplate3 & "<TD align=left valign=center width="& CStr(ControlTabSizes(i)) &" height=20 class='TTableRowUnSelected'><span unselectable='on' DATAFORMATAS=HTML style='overflow:hidden; width:"& CStr(ControlTabSizes(i)-4) &"px;' datafld='"& ControlTabNames(i) &"'></span></TD>" &vbcrlf  
  next
  
  Response.Write DIVTemplate1 & vbCrLf
  Response.Write TableTemplate1
  Response.Write TableTemplate3
  Response.Write TableTemplate2
  Response.Write DIVTemplate2 & vbCrLf
End Sub


' Genereaza DIV-ul principal al controlului in care sunt pozitionate
' apoi celelalte elemente
Private Sub OpenCanvas(strCanvasId, strLeft, strTop, strWidth, strHeight)
  Dim strStyleDef
  Dim strTag
  
  if not(IsNumeric(strLeft) and IsNumeric(strTop)) then
     strStyleDef = "DISPLAY:inline;position:relative;"
  else
     strStyleDef = "POSITION:absolute;LEFT:"+CStr(strLeft)+"px; TOP:"+CStr(strTop)+"px;"
  end if   
  strStyleDef = strStyleDef + " WIDTH:"+CStr(strWidth)+"px; HEIGHT:"+CStr(strHeight)+"px;"
  strStyleDef = strStyleDef +"Z-INDEX:0; background-color:buttonface; OVERFLOW:hidden;"
  
  strTag = "<DIV unselectable='on'"
  strTag = strTag + " id="+strCanvasId
  strTag = strTag + " name="+strCanvasId
  strTag = strTag + " style='"+strStyleDef+"'"
  strTag = strTag + ">" + vbCrLF+vbCrLF
  
  Response.Write strTag
End Sub


' Inchide DIV-ul principal
Private Sub CloseCanvas
  Response.Write (vbCrLf+"</DIV>"+vbCrLf)
End Sub


' Converteste un sir cu spatii intr-un sir de caractere similar
' dar care este corect sintactic pentru a fi folosit pe post de ID
Private Function CvtStrToID(strTabName)
  CvtStrToID = Replace(Replace(Replace(Replace(strTabName,",","_")," ","_"),"(","_"),")","_")
End Function
%>

