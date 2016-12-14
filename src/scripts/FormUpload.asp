<%
Class clsFormItem
 Public FileName
 Public ContentType
 Public Value
 Public FieldName
 Public Length
 Public BinaryData
End Class


Function GetFormItems
 Dim reCount, re()
 Dim ReceivedForm, CursorStart, CursorEnd, PosDisposition, PosFile
 Dim SirDelimitator, SirDelimitatorFinal, SirDelimitatorStart, SirDelimitatorEnd

 reCount = 0
 ReDim re(-1)
		
 ' Citeste binar tot formul receptionat in memorie
 ReceivedForm = Request.BinaryRead(Request.TotalBytes)
		
 ' Obtine sirul delimitator dintre campurile formului
 ' Acesta este de forma: 
 ' -----------------------------7d123e37420264
 ' Precum si sirul care inchide formul de forma:
 ' -----------------------------7d123e37420264--
 CursorStart         = 1
 CursorEnd           = InstrB(CursorStart, ReceivedForm, getByteString(Chr(13)))
 SirDelimitator      = MidB(ReceivedForm, CursorStart, CursorEnd - CursorStart)
 SirDelimitatorFinal = SirDelimitator & getByteString("--")
 SirDelimitatorStart = InstrB(1, ReceivedForm, SirDelimitator)
		
 ' Se bucleaza pana cand se ajunge la sirul delimitator final
 Do Until (SirDelimitatorStart = InstrB(ReceivedForm, SirDelimitatorFinal))
 	ReDim Preserve re(reCount)
 	reCount = reCount + 1
 	Set TmpFrmItem = New clsFormItem

 	PosDisposition        = InstrB(SirDelimitatorStart, ReceivedForm, getByteString("Content-Disposition"))
	CursorStart           = InstrB(PosDisposition, ReceivedForm, getByteString("name=")) + 6
	CursorEnd             = InstrB(CursorStart, ReceivedForm, getByteString(Chr(34)))
	TmpFrmItem.FieldName  = LCase(getString(MidB(ReceivedForm, CursorStart, CursorEnd - CursorStart)))
			
	PosFile               = InstrB(SirDelimitatorStart, ReceivedForm, getByteString("filename="))
	SirDelimitatorEnd     = InstrB(CursorEnd, ReceivedForm, SirDelimitator)
			
	' Daca nu se gaseste sirul "filename=" la pozitia normala inseamna ca
	' respectivul camp este un FormItem normal
	If (PosFile = 0) or (PosFile >= SirDelimitatorEnd) Then
		CursorStart            = InstrB(PosDisposition, ReceivedForm, getByteString(Chr(13))) + 4
		CursorEnd              = InstrB(CursorStart, ReceivedForm, SirDelimitator) - 2
		TmpFrmItem.Value       = getString(MidB(ReceivedForm,CursorStart,CursorEnd-CursorStart))
		TmpFrmItem.Length      = Len(TmpFrmItem.Value)
	Else
		CursorStart            = PosFile + 10
		CursorEnd              = InstrB(CursorStart, ReceivedForm, getByteString(Chr(34)))
		TmpFrmItem.FileName    = getString(MidB(ReceivedForm,CursorStart,CursorEnd-CursorStart))
		
		CursorStart            = InstrB(CursorEnd,ReceivedForm,getByteString("Content-Type:")) + 14
		CursorEnd              = InstrB(CursorStart,ReceivedForm,getByteString(Chr(13)))
		TmpFrmItem.ContentType = getString(MidB(ReceivedForm,CursorStart,CursorEnd-CursorStart))

		CursorStart            = CursorEnd + 4
		CursorEnd              = InstrB(CursorStart,ReceivedForm,SirDelimitator)-2
		Value                  = MidB(ReceivedForm,CursorStart,CursorEnd-CursorStart)
		TmpFrmItem.BinaryData  = Value & getByteString(vbNull)
		TmpFrmItem.Length      = LenB(Value)
	End If

	Set re(UBound(re)) = TmpFrmItem

	SirDelimitatorStart = InstrB(SirDelimitatorStart + LenB(SirDelimitator), ReceivedForm, SirDelimitator)
	Set TmpFrmItem      = Nothing
 Loop
 
 GetFormItems = re
End Function

Function FormItem(fldar, fld)
  For Each it in fldar
   If it.FieldName = LCase(fld) Then Set FormItem = it: Exit Function
  Next
End Function

'String to byte string conversion
Private Function getByteString(StringStr)
 For i = 1 to Len(StringStr)
 	char = Mid(StringStr,i,1)
	getByteString = getByteString & chrB(AscB(char))
 Next
End Function

'Byte string to string conversion
Private Function getString(StringBin)
 getString =""
 For intCount = 1 to LenB(StringBin)
	getString = getString & chr(AscB(MidB(StringBin,intCount,1))) 
 Next
End Function

' Bibliografie: 
' http://www.planet-source-code.com/xq/ASP/txtCodeId.6447/lngWId.4/qx/vb/scripts/ShowCode.htm
%>