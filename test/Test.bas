Attribute VB_Name = "Test"



Public Sub Test()
	Dim format As String: format = "ab}cd{ }ef\{gh{1}ij{ -2 }kl{ \3 }mn{ ""key_1"" }op{{key_2}}qr{ {key_3\}} }st{ :mm""{-dd-""yyyy}uv{""key_4"":}wx{ \5 : mm-dd-yyyy""mm-dd-yyyy""\} }yz{ : }ab""{6:mm-dd-yyyy}"""
	Dim elements() As ParserElement
	
	Dim status As ParsingStatus
	Dim expr As ParserExpression
	status = Parse(format, elements, expr)
	
	Dim e As ParserElement
	Dim out As String: Dim fld As String
	For i = LBound(elements) To UBound(elements)
		e = elements(i)
		
		Select Case e.Kind
		Case ElementKind.elmPlain
			Debug.Print "PLAIN: out(" & i & ")" & VBA.vbTab & "= " & e.Plain
			out = out & e.Plain
		Case ElementKind.elmField
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
	
	Debug.Print
	Debug.Print "STATUS: " & status
	Debug.Print nChr & " characters"
	Debug.Print nElm & " elements"
	Debug.Print "FORMAT: """ & format & """"
	Debug.Print "OUTPUT: """ & out & """"
End Sub
