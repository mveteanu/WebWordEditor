' BEGIN PUBLIC SECTION

' Obtine un sir in format CSV continand ID-urile
' randurilor selectate
Public Function TableGetSelected(tableID)
  Dim tbllin, re
	
  re = ""
  For Each tbllin In tableID.rows 
    If tbllin.cells(0).className = "TTableRowSelected" Then re = re & tbllin.Cells(0).children(0).innerText & ","
  Next
  If Len(re) > 0 Then re = Left(re, Len(re)-Len(","))
  TableGetSelected = re
End Function


' Selecteaza sau deselecteaza toate randurile din tabel
' in functie de parametrul state
Public Sub TableSelectAll(tableID,state)
  Dim tbllin, tblcol
  For Each tbllin In tableID.rows 
    For Each tblcol In tbllin.Cells
      if state = true then
        tblcol.className = "TTableRowSelected" 
      else
        tblcol.className = "TTableRowUnSelected"
      end if
    next  
  next
End Sub


' BEGIN PRIVATE SECTION


Private Sub HandleTableClick(tipsel)
	Dim obj, sTag, tbllin, mytbl

    Set tbllin  = nothing
    Set mytbl = nothing
   
	Set obj=window.event.srcElement
	sTag = obj.tagName
	Set tbllin = Nothing
	If sTag = "TD" Then 
	  set tbllin  = obj.parentElement
	  set mytbl = obj.parentElement.parentElement.parentElement
	ElseIf sTag = "SPAN" Then
	  set tbllin  = obj.parentElement.parentElement
	  set mytbl = obj.parentElement.parentElement.parentElement.parentElement
	End If
	
    If Not mytbl Is Nothing Then TableSwitchRow mytbl, tbllin , tipsel
End Sub


Private Sub TableSwitchRow(objTabl, objRow, tipsel)
 Dim ir,ic
 
 If tipsel=1 then
   If objRow.Cells(0).className = "TTableRowSelected" Then Exit Sub
   For Each ir In objTabl.rows
     If ir.Cells(0).className = "TTableRowSelected" Then
       For Each ic In ir.cells
         ic.className = "TTableRowUnSelected"
       Next
     End If
   Next
   For Each ic In objRow.cells
     ic.className = "TTableRowSelected"
   Next
 Else
 	For Each ic In objRow.cells
		If ic.className = "TTableRowSelected" Then
			ic.className = "TTableRowUnSelected"
		ElseIf ic.ClassName = "TTableRowUnSelected" Then
			ic.className = "TTableRowSelected"
		End If
	Next
 End If	
End Sub


Private Sub HandleTableHeaderSort(tdcname, nrdiv, lastdiv)
 Dim obj, sTag, maindiv, sortk, sortname, i

 Set obj=window.event.srcElement
 sTag = obj.tagName

 if sTag="DIV" then
   sortk = obj.children(0).innerhtml
   sortname = obj.children(1).innerhtml
   set maindiv = obj.parentElement
 elseif sTag = "SPAN" Then
   sortk = obj.parentElement.children(0).innerhtml
   sortname = obj.parentElement.children(1).innerhtml
   set maindiv = obj.parentElement.parentElement
 end if
 for i=0 to lastdiv
   if i<>nrdiv then
     maindiv.children(i).children(0).style.display="none"
   else
     maindiv.children(i).children(0).innerhtml=CStr(CInt(sortk) xor 3)
     maindiv.children(i).children(0).style.display=""
   end if
 next 

 if sortk="6" then 
    sortk="+"
 else
    sortk="-"
 end if

 tdcname.sort=sortk & sortname
 tdcname.reset
End Sub
