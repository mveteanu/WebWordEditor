<%
' Adauga de pe server un control de tip TDC intr-o pagina
' prin folosirea tag-ului OBJECT
Public Sub AddTDC(strID, strDelim, strURL)
 With Response
   .Write "<OBJECT id='"& strID &"' CLASSID='clsid:333C7BC4-460F-11D0-BC04-0080C7055A83' VIEWASTEXT>" & vbCrLf
   .Write "<PARAM NAME='UseHeader' VALUE='True'>" & vbCrLf
   .Write "<PARAM NAME='FieldDelim' VALUE='"& strDelim &"'>" & vbCrLf
   .Write "<PARAM NAME='DataURL' VALUE='"& strURL &"'>" & vbCrLf
   .Write "</OBJECT>" & vbCrLf & vbCrLf
 End With
End Sub


' Intoarce codul HTML corespunzator pentru umplerea unui SELECT
' folosind taguri <OPTION> folosind un obiect Dictionary
' SelectID indica optiunea SELECTED
Function GetFillSelectFromDict(Dict, SelectID)
 Const OptionMach = "<OPTION @3 value='@1'>@2</OPTION>"
 Dim re, re1, it

 For each it in Dict.keys
  If CStr(SelectID) = CStr(it) then
    re1 = Replace(OptionMach, "@3", "SELECTED")
  Else
    re1 = Replace(OptionMach, "@3", "")  
  End If  
  re  = re & Replace(Replace(re1, "@2", Dict.item(it)), "@1", it) & vbCrLf
 Next
 
 GetFillSelectFromDict = re
End Function
%>

