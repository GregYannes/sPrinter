Attribute VB_Name = "Test"



Public Sub Test()
	Dim format As String: format = "ab}cd{}ef\{gh{1}ij{-2}kl{\3}mn{""key_1""}op{{key_2}}qr{{key_3\}}}st{:mm""{-dd-""yyyy}uv{""key_4"":}wx{\5:mm-dd-yyyy""mm-dd-yyyy""\}}yz""{6:mm-dd-yyyy}"""
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
				
				' If VBA.VarType(e.Field.Index) = VBA.VbVarType.vbString Then
				' 	fld = fld & "'" & e.Field.Index & "'"
				' Else
				' 	fld = fld & e.Field.Index
				' End If
			End If
			
			If e.Field.Format <> VBA.vbNullString Then
				fld = fld & ":" & e.Field.Format
			End If
			
			fld = fld & "}"
			Debug.Print "FIELD: out(" & i & ")" & VBA.vbTab & "= " & fld
			
			out = out & fld
		End Select
	Next i
	
	Debug.Print
	Debug.Print VBA.Len(format) & " chars"
	Debug.Print "status = " & status
	Debug.Print "lower = " & LBound(elements) & ", upper = " & UBound(elements)
	Debug.Print "OUTPUT: """ & out & """"
	Debug.Print "FORMAT: """ & format & """"
End Sub
