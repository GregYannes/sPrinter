Attribute VB_Name = "Test"



' Test all relevant features of "sPrinter".
Public Sub Test()
	' Parse a message.
	Test_Parse
	
	Debug.Print
	Debug.Print
	Debug.Print
	Debug.Print
	
	Debug.Print
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
	Const DEFAULT_VALUE As Variant = "?"
	
	
	Dim fmt0 As String: fmt0 = "You have a meeting with {0} {1} at {2:h:MM AM/PM} on {3:dddd, mmmm d}."
	Dim fmt1 As String: fmt1 = "You have a meeting with {1} {2} at {3:h:MM AM/PM} on {4:dddd, mmmm d}."
	Dim fmt2 As String: fmt2 = "You have a meeting with {1} {2} at {-2:h:MM AM/PM} on {-1:dddd, mmmm d}."
	Dim fmt3 As String: fmt3 = "You have a meeting with {} {} at {:h:MM AM/PM} on {:dddd, mmmm d}."
	Dim fmt4 As String: fmt4 = "You have a meeting with {{forename}} {{surname}} at {{time}:h:MM AM/PM} on {{date}:dddd, mmmm d}."
	Dim fmt5 As String: fmt5 = "You have a meeting with {} {2} at {-2:h:MM AM/PM} on {{date}:dddd, mmmm d} regarding {5} and {-5} and {{omitted}}."
	
	
	Dim arrData As Variant: arrData = Array( _
		"John", _
		"Doe", _
		VBA.Time(), _
		VBA.Date() _
	)
	
	Dim arrLookup As Variant: arrLookup = Array( _
		"Forename", _
		"Surname", _
		"Time", _
		"Date" _
	)
	
	
	Dim clxData As Collection: Set clxData = New Collection
	Dim i As Long
	For i = LBound(arrData) To UBound(arrData)
		clxData.Add arrData(i), key := arrLookup(i)
	Next i
	
	
	Dim arrMsg0 As String: arrMsg0 = sPrinter.Message(fmt0, arrData,          , default := DEFAULT_VALUE)
	Dim arrMsg2 As String: arrMsg2 = sPrinter.Message(fmt2, arrData,          , default := DEFAULT_VALUE, position := posRelative)
	Dim arrMsg3 As String: arrMsg3 = sPrinter.Message(fmt3, arrData,          , default := DEFAULT_VALUE)
	Dim arrMsg4 As String: arrMsg4 = sPrinter.Message(fmt4, arrData, arrLookup, default := DEFAULT_VALUE)
	Dim arrMsg5 As String: arrMsg5 = sPrinter.Message(fmt5, arrData, arrLookup, default := DEFAULT_VALUE, position := posRelative)
	
	
	Dim clxMsg1 As String: clxMsg1 = sPrinter.Message(fmt1, clxData,          , default := DEFAULT_VALUE)
	Dim clxMsg2 As String: clxMsg2 = sPrinter.Message(fmt2, clxData,          , default := DEFAULT_VALUE, position := posRelative)
	Dim clxMsg3 As String: clxMsg3 = sPrinter.Message(fmt3, clxData,          , default := DEFAULT_VALUE)
	Dim clxMsg4 As String: clxMsg4 = sPrinter.Message(fmt4, clxData,          , default := DEFAULT_VALUE)
	Dim clxMsg5 As String: clxMsg5 = sPrinter.Message(fmt5, clxData, arrLookup, default := DEFAULT_VALUE, position := posRelative)
	
	
	Debug.Print "INPUT:" & VBA.vbTab & fmt0
	Debug.Print "OUTPUT:" & VBA.vbTab & arrMsg0
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt2
	Debug.Print "OUTPUT:" & VBA.vbTab & arrMsg2
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt3
	Debug.Print "OUTPUT:" & VBA.vbTab & arrMsg3
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt4
	Debug.Print "OUTPUT:" & VBA.vbTab & arrMsg4
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt5
	Debug.Print "OUTPUT:" & VBA.vbTab & arrMsg5
	
	Debug.Print
	Debug.Print
	Debug.Print
	Debug.Print
	
	Debug.Print "INPUT:" & VBA.vbTab & fmt1
	Debug.Print "OUTPUT:" & VBA.vbTab & clxMsg1
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt2
	Debug.Print "OUTPUT:" & VBA.vbTab & clxMsg2
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt3
	Debug.Print "OUTPUT:" & VBA.vbTab & clxMsg3
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt4
	Debug.Print "OUTPUT:" & VBA.vbTab & clxMsg4
	Debug.Print
	Debug.Print
	Debug.Print "INPUT:" & VBA.vbTab & fmt5
	Debug.Print "OUTPUT:" & VBA.vbTab & clxMsg5
End Sub
