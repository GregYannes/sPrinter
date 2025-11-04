Attribute VB_Name = "Test"



Public Sub Test()
	Dim format As String: format = "ab}cd{}ef\{gh{1}ij{-2}kl{\3}mn{""key_1""}op{{key_2}}qr{{key_3\}}}st{:mm""{-dd-""yyyy}uv{""key_4"":}wx{\5:mm-dd-yyyy""mm-dd-yyyy""\}}yz""{6:mm-dd-yyyy}"""
	Dim elements() As sParseElement
	
	Dim status As sParseStatus
	status = sParse(format, elements)
	
	Dim e As sParseElement
	Dim out As String: Dim fld As String
	For i = LBound(elements) To UBound(elements)
		e = elements(i)
		' Debug.Print e.Kind
		
		Select Case e.Kind
		Case sParseKind.pkPlain
			Debug.Print "PLAIN: out(" & i & ") = '" & e.Text & "'"
			out = out & e.Text
		Case sParseKind.pkField
			fld = "{"
			
			If e.HasIndex Then
				fld = fld & e.RawIndex
				' If e.IndexIsKey Then
				' 	fld = fld & "'" & e.Index & "'"
				' Else
				' 	fld = fld & e.Index
				' End If
			End If
			
			If e.HasFormat Then
				fld = fld & ":" & e.format
			End If
			
			fld = fld & "}"
			Debug.Print "FIELD: out(" & i & ") = " & fld
			
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
