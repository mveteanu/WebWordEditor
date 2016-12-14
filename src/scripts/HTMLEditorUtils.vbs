Dim FileNumber
Dim DSel
FileNumber = 0

' Intoarce true daca valoarea v se afla printre elementele
' array-ului ar
Function InArray(v,ar)
 Dim re

 re = false
 
 If Join(ar)<>"" then
	for i=0 to UBound(ar)
	 if ar(i) = v then
	   re = true
	   exit for
	 end if
	next
 End If
 
 InArray = re
End Function

' Intoarce un array ce reprezinta diferenta array-urilor
' biga si littlea, adica un array format din acele elemente 
' care se afla in biga si nu se gasesc in littlea
Function ArrayDif(biga, littlea)
 Dim re()
 Dim nri
 
 If Join(biga)<>"" then
	nri = 0
	for i=0 to UBound(biga)
	 If not InArray(biga(i),littlea) then
	   Redim Preserve re(nri)
	   re(nri) = biga(i)
	   nri = nri + 1
	 end if
	next
 End If
 
 ArrayDif = re
End Function


' Decodeaza informatia dintr-un URL
' prin inlocuirea secventelor de tipul %xx cu caracterele corespunzatoare
' Nu se traduce caracterul + in " " datorita modului in care IE trateaza acest caracter
' Nu se trateaza secventele de tipul %uxxxx !
Function URLDecode(Expression)
  Dim strSource, strTemp, strResult
  Dim lngPos, s
  strSource = Expression
  For lngPos = 1 To Len(strSource)
    strTemp = Mid(strSource, lngPos, 1)
    If strTemp = "%" Then
      If lngPos + 2 < Len(strSource) Then
        s = Mid(strSource, lngPos + 1, 2)
        If IsNumeric("&H" & s) then 
           strResult = strResult & Chr(CInt("&H" & s))
        Else
           strResult = strResult & "%" & s
        End If   
        lngPos = lngPos + 2
      End If
    Else
      strResult = strResult & strTemp
    End If
  Next
  URLDecode = strResult
End Function

' Obtine valoarea unui parametru encodat intr-un URL
' Are un efect asemanator cu metoda ASP: Server.RequestQueryString
Function QueryString(url, item)
 Dim p1, p2
 Dim re
 
 p1 = InstrRev(url,item & "=") + Len(item & "=")
 if p1 = Len(item & "=") then 
   re = ""
 else  
   p2 = Instr(p1,url,"&")
   if p2 = 0 then p2 = Len(url) + 1
   re = Mid(url,p1,p2-p1)
 end if 
  
 QueryString = re
End Function


' Intoarce un Array cu SRC-ul imaginilor dintr-un container
' specificat prin EditPage
Function GetEditPageImages(EditPage)
 Dim im()
 Dim nri
 
 nri = 0
 for each i in EditPage.All
   if i.tagName = "IMG" then 
     Redim Preserve im(nri)
     im(nri) = i.src
     nri = nri + 1
   end if
 next
 
 GetEditPageImages = im
End Function


' Intoarce un Array cu SRC-ul imaginilor dintr-un container
' specificat prin EditPage care vin de pe serverul de web prin 
' protocol http:// (nu se verifica de pe ce server vin imaginile !!!)
Function GetEditPageServerImages(EditPage)
 Dim im()
 Dim nri
 
 nri = 0
 for each i in EditPage.All
   if i.tagName = "IMG" then
     if LCase(Left(i.src,7)) = "http://" then 
       Redim Preserve im(nri)
       im(nri) = i.src
       nri = nri + 1
     end if
   end if  
 next
 
 GetEditPageServerImages = im
End Function



' Intoarce un Array cu SRC-ul imaginilor dintr-un container
' specificat prin EditPage aflate local (pe HD sau pe disk sharat) 
Function GetEditPageLocalImages(EditPage)
 Dim im()
 Dim nri
 
 nri = 0
 for each i in EditPage.All
   if i.tagName = "IMG" then
     if LCase(Left(i.src,7)) = "file://" then 
       Redim Preserve im(nri)
       im(nri) = i.src
       nri = nri + 1
     end if
   end if  
 next
 
 GetEditPageLocalImages = im
End Function


' Converteste un string ce reprezinta url-ul unui fisier local
' din formatele: file:///d:/t/img/... sau file://vma-athlon/t/img/...
' in formatele:  d:\t\img\... sau \\vma-athlon\t\img\...
' Daca sirul de intrare nu respecta conventiile amintite atunci
' se intoarce nemodificat
Function LocalURLToFileName(fileurl)
 Dim s
 Dim re
 
 s = Replace(fileurl,"/","\")
 if LCase(Left(s,8))="file:\\\" then 
   re = URLDecode(Right(s,Len(s)-8))
 elseif LCase(Left(s,7))="file:\\" then
   re = URLDecode("\\" & Right(s,Len(s)-7))
 elseif (UCase(Left(s,1))>="A") and (UCase(Left(s,1))<="Z") and (Mid(s,2,2)=":\") then
   re = URLDecode(s)
 else
   re = fileurl
 end if
 
 LocalURLToFileName = re
End Function


' Extrage valorile specificate prin idname dintr-un array de URL-uri
' si le concateneaza intr-un string in format CSV
Function FileNamesArrayToIDCSV(ar, idname)
 Dim re
 Dim qs
 
 re = ""
 If Join(ar)<>"" then
  For i = 0 to UBound(ar)
   qs = QueryString(ar(i),idname)
   If qs<>"" then re = re & qs & ","
  Next
  re = Left(re,Len(re)-1)
 End If
 
 FileNamesArrayToIDCSV = re
End Function


' Se executa la apasarea butonului Insert Picture si are ca efect 
' afisarea miniformului pentru introducerea numelui fisierului
Sub ShowAddFileForm(ByRef FileNr)
 set DSel = document.selection.createRange 
 FileNr = FileNr + 1
 ChosePictureFormular.insertAdjacentHTML "beforeEnd",  "<INPUT type='file' name='File" & CStr(FileNr) & "' id='File" & CStr(FileNr) & "'>"
 set ifile = ChosePictureForm.all("File" & CStr(FileNr))
 with ifile
  .className = "TEdit"
  .style.left = 20
  .style.top = 40
  .style.width = 250
  .style.height = 21
  ChosePictureForm.style.visibility = ""
  .focus
 end with
 set ifile = nothing
End Sub


' Se executa automat la apasarea butonului OK din cadrul miniformului
' creat de subrutina anterioara (cea care se executa la Insert Picture)
Sub ChosePictureFormButOK_onclick
 Dim im 
 im = ChosePictureForm.all("File" & CStr(FileNumber)).Value
 if im<>"" then
   if LCase(document.selection.type) <> "control" then 
     DSel.select
     document.execCommand "InsertImage", false, im
   end if  
 end if
End Sub


' Sterge din formul cu fisiere ascuns acele elemente INPUT FILE 
' a caror imagini indicate au fost sterse din documentul ce se editeaza
' sau cele care apar in mai multe exemplare
Sub CleanFilesForm(fform, docimg)
 Dim este, fty
 
 for each ff in fform.elements
  if ff.type = "file" then
	este = false
	If Join(docimg)<>"" then
		for li = 0 to UBound(docimg)
		  if ff.value = LocalURLToFileName(docimg(li)) then este = true
		next 
	End If 	
	if not este then fform.removeChild ff
  end if
 next
 ' Acum se vor elimina imaginile duplicat pentru
 ' micsorarea traficului pe retea si a spatiului din BD
 for i1 = 0 to fform.elements.length-1
   for i2 = i1+1 to fform.elements.length-1
     if (LCase(fform.elements(i1).type)="file") and (LCase(fform.elements(i2).type)="file") and (fform.elements(i1).value = fform.elements(i2).value) then
       fform.removeChild fform.elements(i1)
       Exit For
     end if
   next
 next
End Sub


' Obtine sub forma de CSV numele campurilor din formul ascuns
' ce contin fisierele cu imagini ce trebuie uploadate
Function GetUploadFields(fform)
 Dim re
 
 re = ""
 For each ff in fform.elements
  if ff.type = "file" then
    re = re & CStr(ff.name) & ","
  end if  
 Next
 If Len(re)>0 then re = Left(re,Len(re)-1)
 
 GetUploadFields = re
End Function
