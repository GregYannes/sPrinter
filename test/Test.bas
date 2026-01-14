Attribute VB_Name = "Test"



' Test all relevant features of "sPrinter".
Public Sub Test()
	' Parse a message.
	Test_Parse
	
	Debug.Print
	Debug.Print
	Debug.Print
	
	' Format and display another message.
	Test_Message
End Sub


' Parse a message.
Public Sub Test_Parse()
	Dim format As String: format = "ab}cd{ }ef\{gh{1}ij{ -2 }kl{ \3 }mn{ ""key_1"" }op{{key_2}}qr{ {key_3\}} }st{ :mm""{-dd-""yyyy}uv{""key_4"":}wx{ \5 : mm-dd-yyyy""mm-dd-yyyy""\} }yz{ : }ab""{6:mm-dd-yyyy}"""
	Dim elements() As sPrinter.ParserElement
	elements = sPrinter.Parse(format := format)
	
	Dim e As sPrinter.ParserElement
	Dim out As String: Dim fld As String
	For i = LBound(elements) To UBound(elements)
		e = elements(i)
		
		Select Case e.Kind
		Case sPrinter.ElementKind.elmPlain
			Debug.Print "PLAIN: out(" & i & ")" & VBA.vbTab & "= " & e.Plain
			out = out & e.Plain
		Case sPrinter.ElementKind.elmField
			fld = "{"
			
			If Not VBA.IsEmpty(e.Field.Index) Then
				If VBA.VarType(e.Field.Index) = VBA.VbVarType.vbString Then
					fld = fld & "'" & e.Field.Index & "'"
				Else
					fld = fld & e.Field.Index
				End If
			End If
			
			If e.Field.Format <> VBA.vbNullString Then
				fld = fld & ":" & e.Field.Format
			End If
			
			fld = fld & "}"
			Debug.Print "FIELD: out(" & i & ")" & VBA.vbTab & "= " & fld
			
			out = out & fld
		Case Else
			Debug.Print "UNKNOWN: out(" & i & ")"
		End Select
	Next i
	
	Dim nChr As Long: nChr = VBA.Len(format)
	Dim lElm As Long: lElm = LBound(elements)
	Dim uElm As Long: uElm = UBound(elements)
	Dim nElm As Long: nElm = uElm - lElm + 1
	' Dim nElm As Long: nElm = sPrinter.Elm_Count(elements)
	
	Debug.Print
	Debug.Print nChr & " characters"
	Debug.Print nElm & " elements"
	Debug.Print "FORMAT:" & VBA.vbTab & format
	Debug.Print "OUTPUT:" & VBA.vbTab & out
End Sub


' Format and display a message.
Public Sub Test_Message()
	Dim fmt1 As String: fmt1 = "You have a meeting with {1} {2} at {3:h:MM AM/PM} on {4:dddd, mmmm d}."
	Dim fmt2 As String: fmt2 = "You have a meeting with {1} {2} at {-2} on {-1}."
	Dim fmt3 As String: fmt3 = "You have a meeting with {} {} at {:h:MM AM/PM} on {:dddd, mmmm d}."
	Dim fmt4 As String: fmt4 = "You have a meeting with {{forename}} {{surname}} at {{time}:h:MM AM/PM} on {{date}:dddd, mmmm d}."
	
	Dim data As Collection: Set data = New Collection
	data.Add "John",	key := "Forename"
	data.Add "Doe",		key := "Surname"
	data.Add VBA.Time(),	key := "Time"
	data.Add VBA.Date(),	key := "Date"
	
	
	Dim msg1 As String: msg1 = sPrinter.Message(fmt1, data, default := "?")
	Dim msg2 As String: msg2 = sPrinter.Message(fmt2, data, default := "?", position := posRelative)
	Dim msg3 As String: msg3 = sPrinter.Message(fmt3, data, default := "?")
	Dim msg4 As String: msg4 = sPrinter.Message(fmt4, data, default := "?")
	
	
	Debug.Print "INPUT:" & VBA.vbTab & fmt1
	Debug.Print "OUTPUT:" & VBA.vbTab & msg1
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt2
	Debug.Print "OUTPUT:" & VBA.vbTab & msg2
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt3
	Debug.Print "OUTPUT:" & VBA.vbTab & msg3
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt4
	Debug.Print "OUTPUT:" & VBA.vbTab & msg4
End Sub
