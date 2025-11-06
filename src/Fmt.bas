Attribute VB_Name = "Fmt"



' #############
' ## Options ##
' #############

' Explicitly declare all variables.
Option Explicit


' ' Hide these functions from Excel.
' Option Private Module



' ##############
' ## Metadata ##
' ##############

Public Const MOD_NAME As String = "Fmt"

Public Const MOD_VERSION As String = ""

Public Const MOD_REPO As String = "https://github.com/GregYannes/Fmt"



' ###############
' ## Constants ##
' ###############

' Syntax for parsing.
Private Const STX_ESC As String = "\"			' Escape the next character.
Private Const STX_FLD_OPEN As String = "{"		' Embed a field for formatting...
Private Const STX_FLD_CLOSE As String = "}"		' ...and enclose that field.
Private Const STX_QUO_OPEN As String = """"		' Quote the next several characters...
Private Const STX_QUO_CLOSE As String = STX_QUO_OPEN	' ...and enclose that quote.
Private Const STX_SEP As String = ":"			' Separate specifiers in a field.



' ##################
' ## Enumerations ##
' ##################

' Engine used for formatting.
Public Enum FormatMode
	fmVbFormat	' The Format() function in VBA.
	fmXlText	' The Text() function in Excel.
End Enum


' Outcomes of parsing.
Public Enum ParsingStatus
	psSuccess = 0			' Report success.
	psError = 1000			' Report a general syntax error.
	psErrorHangingEscape = 1001	' Report a hanging escape...
	psErrorUnclosedField = 1002	' ...or an incomplete field...
	psErrorUnclosedQuote = 1003	' ...or an incomplete quote...
	psErrorNonintegralIndex = 1004	' ...or an index that is not an integer.
End Enum


' Kinds of elements which may be parsed.
Public Enum ElementKind
	[_Unknown]	' Uninitialized.
	ekPlain		' Plain text which is displayed as is.
	ekField		' Field that is formatted and embedded.
End Enum


' Contexts in which symbols are interpreted.
Private Enum ParsingContext
	[_Unknown]	' Uninitialized.
	pcPlain		' Plain text.
	pcField		' An embedded field...
	pcFieldIndex	' ...its index...
	pcFieldFormat	' ...and its format.
End Enum


' Ways to defuse literal symbols rather than interpreting them.
Private Enum ParsingDefusal
	[_Off]		' No defusal.
	pdEscape	' Defuse only the next character...
	pdQuote		' ...or all characters within quotes.
End Enum


' Kinds of indices for extracting values.
Private Enum IndexKind
	[_Unknown]	' Uninitialized.
	ikPosition	' Integer for a position...
	ikKey		' ...or text for a key.
End Enum



' ###########
' ## Types ##
' ###########

' Element for parsing the index...
Public Type peFieldIndex
	Exists As Boolean
	Kind As IndexKind
	Code As String
	Start As Long
	Stop As Long
	Position As Long
	Key As String
End Type


' ...and the custom format...
Public Type peFieldFormat
	Exists As Boolean
	Start As Long
	Stop As Long
	Code As String
End Type


' ...of a field embedded in formatting.
Public Type peField
	Code As String
	Start As Long
	Stop As Long
	Index As peFieldIndex
	Format As peFieldFormat
End Type


' Element for parsing plain text in formatting.
Public Type pePlain
	Text As String
End Type


' Elements into which formats are parsed.
Public Type ParsingElement
	Kind As ElementKind
	Code As String
	Start As Long
	Stop As Long
	Plain as pePlain
	Field As peField
End Type



' #########
' ## API ##
' #########

' .
Public Function Parse( _
	ByRef format As String, _
	ByRef elements() As ParsingElement, _
	Optional ByRef charIndex As Long, _
	Optional ByVal base As Long = 1, _
	Optional ByVal escape As String = STX_ESC, _
	Optional ByVal openField As String = STX_FLD_OPEN, _
	Optional ByVal closeField As String = STX_FLD_CLOSE, _
	Optional ByVal openQuote As String = STX_QUO_OPEN, _
	Optional ByVal closeQuote As String = STX_QUO_CLOSE, _
	Optional ByVal separator As String = STX_SEP _
) As ParsingStatus
	
	' ###########
	' ## Setup ##
	' ###########
	
	' Record the format length.
	Dim fLen As Long: fLen = VBA.Len(format)
	
	' Short-circuit for unformatted input.
	If fLen = 0 Then
		charIndex = 0
		Erase elements
		Parse = ParsingStatus.psSuccess
		Exit Function
	End If
	
	
	' Size to accommodate all (possible) elements.
	Dim eLen As Long: eLen = VBA.Int(fLen / 2) + 1
	Dim eUp As Long: eUp = base + eLen - 1
	ReDim elements(base To eUp)
	
	
	' Track the current context for parsing...
	Dim cxt As ParsingContext: cxt = ParsingContext.[_Unknown]
	Dim isQuo As Boolean: isQuo = False
	Dim isEsc As Boolean: isEsc = False
	
	' ...and the current depth of nesting...
	Dim depth As Long: depth = 0
	
	' ...and the current element...
	Dim eIdx As Long: eIdx = base
	
	' ...and the current characters.
	Dim char As String: charIndex = 1
	Dim nQuo As Long: nQuo = 0
	Dim idxEsc As Boolean: idxEsc = False
	Dim idxStart As Long, idxStop As Long, idxLen As Long
	Dim fmtStart As Long, fmtStop As Long, fmtLen As Long
	Dim fldStatus As ParsingStatus: fldStatus = ParsingStatus.psSuccess
	
	
	
	' #############
	' ## Parsing ##
	' #############
	
	' Catch generic syntax errors.
	On Error GoTo STX_ERROR
	
	' Scan and parse the format.
	Do While charIndex <= fLen
		
		' Extract the current character.
		char = VBA.Mid$(format, charIndex, 1)
		
		' Interpret this character in context.
		Select Case cxt
		
		
		
		' ##############
		' ## Inactive ##
		' ##############
		
		Case ParsingContext.[_Unknown]
			Select Case char
			
			' Parse into a field...
			Case openField
				depth = depth + 1
				cxt = ParsingContext.pcField
				elements(eIdx).Kind = ElementKind.ekField
				GoTo NEXT_CHAR
				
			' ...or interpret as text.
			Case Else
				cxt = ParsingContext.pcPlain
				elements(eIdx).Kind = ElementKind.ekPlain
				GoTo NEXT_LOOP
			End Select
			
			
			
		' ################
		' ## Plain Text ##
		' ################
		
		Case ParsingContext.pcPlain
			
			' Quote "inert" text...
			If isQuo Then
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					isQuo = False
					
				' ...or continue quoting.
				Case Else
					elements(eIdx).Text = elements(eIdx).Text & char
				End Select
				
			' ...or escape literal text...
			ElseIf isEsc Then
				elements(eIdx).Text = elements(eIdx).Text & char
				isEsc = False
				
			' ...or parse "active" text.
			Else
				Select Case char
				
				' Quote the next characters...
				Case openQuote
					isQuo = True
					
				' ..escape the next character...
				Case escape
					isEsc = True
					
				' ...or parse into a field...
				Case openField
					' Update parsing context.
					depth = depth + 1
					cxt = ParsingContext.pcField
					
					' Move to the next element if the current is already used.
					If elements(eIdx).Kind <> ElementKind.[_Unknown] Then
						eIdx = eIdx + 1
					End If
					
					' Identify the element as a field.
					elements(eIdx).Kind = ElementKind.ekField
					
				' ...or display literally.
				Case Else
					elements(eIdx).Text = elements(eIdx).Text & char
				End Select
			End If
			
			GoTo NEXT_CHAR
			
			
			
		' ###########
		' ## Field ##
		' ###########
		
		Case ParsingContext.pcField
			Select Case char
			
			' Parse out of the field...
			Case closeField
				depth = depth - 1
				If depth = 0 Then GoTo END_FIELD
				
			' ...or parse into the format...
			Case separator
				cxt = ParsingContext.pcFieldFormat
				elements(eIdx).HasFormat = True
				fmtStart = charIndex
				fmtStop = fmtStart
				
			' ...or parse the index.
			Case Else
				cxt = ParsingContext.pcFieldIndex
				elements(eIdx).HasIndex = True
				idxStart = charIndex
				idxStop = idxStart
				
				GoTo NEXT_LOOP
			End Select
			
			GoTo NEXT_CHAR
			
			
			
		' ###################
		' ## Field | Index ##
		' ###################
		
		Case ParsingContext.pcFieldIndex
			
			' Quote "inert" symbol...
			If isQuo Then
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					isQuo = False
					If depth = 1 Then cxt = ParsingContext.pcField
					
				' ...or continue quoting.
				Case Else
					elements(eIdx).Index = elements(eIdx).Index & char
				End Select
				
			' ...or escape literal symbol...
			ElseIf isEsc Then
				elements(eIdx).Index = elements(eIdx).Index & char
				isEsc = False
				If depth = 1 Then cxt = ParsingContext.pcField
				
			' ...or parse "active" symbol.
			Else
				Select Case char
				
				' Escape the next character...
				Case escape
					isEsc = True
					idxEsc = True
					
				' ...or nest into the field...
				Case openField
					depth = depth + 1
					If depth = 2 Then
						nQuo = nQuo + 1
					Else
						elements(eIdx).Index = elements(eIdx).Index & char
					End If
					
				' ...or unnest out of the field...
				Case closeField
					depth = depth - 1
					If depth = 0 Then
						cxt = ParsingContext.[_Unknown]
						GoTo END_FIELD
					ElseIf depth = 1 Then
						cxt = ParsingContext.pcField
					Else
						elements(eIdx).Index = elements(eIdx).Index & char
					End If
					
				' ...or parse into a quoted key...
				Case openQuote
					isQuo = True
					If depth = 1 Then nQuo = nQuo + 1
					
				' ' ...or parse into a format...
				' Case separator
				' 	cxt = ParsingContext.pcFormat
				' 	elements(eIdx).HasFormat = True
					
				' ...or display literally.
				Case Else
					elements(eIdx).Index = elements(eIdx).Index & char
				End Select
			End If
			
			idxStop = idxStop + 1
			GoTo NEXT_CHAR
			
			
			
		' ####################
		' ## Field | Format ##
		' ####################
		
		Case ParsingContext.pcFieldFormat
			
			' Include quoted symbol...
			If isQuo Then
				' Terminate the quote if appropriate.
				If char = closeQuote Then isQuo = False
				
			' ...or include escaped symbol...
			ElseIf isEsc Then
				isEsc = False
				
			' ...but parse "active" symbol.
			Else
				Select Case char
				
				' Escape the next character...
				Case escape
					isEsc = True
					
				' ...or nest into the field...
				Case openField
					depth = depth + 1
					
				' ...or unnest out of the field...
				Case closeField
					depth = depth - 1
					If depth = 0 Then GoTo END_FIELD
					
				' ...or parse into a quoted key.
				Case openQuote
					isQuo = True
				End Select
				
				
			End If
			
			fmtStop = fmtStop + 1
			GoTo NEXT_CHAR
			
		End Select
		
		
		
	' #############
	' ## Control ##
	' #############
	
	' ' Save the information to the element.
	' SAVE_ELEMENT:
	' 	elements(eIdx).Kind = e_Kind
	' 	elements(eIdx).Text = e_Text
	' 	elements(eIdx).HasIndex = e_HasIndex
	' 	elements(eIdx).Index = e_Index
	' 	elements(eIdx).IndexRaw = e_IndexRaw
	' '	elements(eIdx).IndexIsKey = e_IndexIsKey
	' '	elements(eIdx).EscapesIndex = e_EscapesIndex
	' '	elements(eIdx).QuotesIndex = e_QuotesIndex
	' 	elements(eIdx).HasFormat = e_HasFormat
	' 	elements(eIdx).Format = e_Format
	' 	
	' ' Reset the information.
	' RESET_ELEMENT:
	' 	e_Kind = ElementKind.[_Unknown]
	' 	e_Text = VBA.vbNullString
	' 	e_HasIndex = False
	' 	e_Index = VBA.vbNullString
	' 	e_IndexRaw = VBA.vbNullString
	' 	e_IndexIsKey = False
	' '	e_EscapesIndex = False
	' ' 	e_QuotesIndex = False
	' 	e_HasFormat = False
	' 	e_Format = VBA.vbNullString
		
	' Parse out of the field.
	END_FIELD:
		' Record the elemental information...
		fldStatus = EndField( _
			format := format, _
			e := elements(eIdx), _
			cxt := cxt, _
			nQuo := nQuo, _
			idxEsc := idxEsc, _
			idxStart := idxStart, _
			idxStop := idxStop, _
			fmtStart := fmtStart, _
			fmtStop := fmtStop _
		)
		
		' ...and short-circuit for an index of the wrong type.
		If fldStatus = ParsingStatus.psErrorNonintegralIndex Then Exit Do
		fldStatus = ParsingStatus.psSuccess
		
		' Increment the element.
		eIdx = eIdx + 1
		
		GoTo NEXT_CHAR
		
	' ' Increment the element.
	' NEXT_ELEMENT:
	' 	eIdx = eIdx + 1
		
	' Increment the character.
	NEXT_CHAR:
		charIndex = charIndex + 1
		
	' Continue to the next iteration.
	NEXT_LOOP:
		
	Loop
	
	
	' Deactivate error handling.
	On Error GoTo 0
	
	
	
	' ####################
	' ## Interpretation ##
	' ####################
	
	' Resize to the elements we actually parsed.
	If elements(eIdx).Kind = ElementKind.[_Unknown] Then
		eIdx = eIdx - 1
	End If
	
	If eIdx < eUp Then
		eUp = eIdx
		ReDim Preserve elements(base To eUp)
	End If
	
	
	' Record any pending field information.
	Select Case cxt
	Case ParsingContext.pcField, ParsingContext.pcFieldIndex, ParsingContext.pcFieldFormat
		fldStatus = EndField( _
			format := format, _
			e := elements(eIdx), _
			cxt := cxt, _
			nQuo := nQuo, _
			idxEsc := idxEsc, _
			idxStart := idxStart, _
			idxStop := idxStop, _
			fmtStart := fmtStart, _
			fmtStop := fmtStop _
		)
	End Select
	
	
	' Report status: a hanging escape...
	If isEsc Then
		Parse = ParsingStatus.psErrorHangingEscape
		
	' ...or an unclosed quote...
	ElseIf isQuo Then
		Parse = ParsingStatus.psErrorUnclosedQuote
		
	' ...or an unclosed field...
	ElseIf depth <> 0 Then
		Parse = ParsingStatus.psErrorUnclosedField
		
	' ...or a index of the wrong type...
	ElseIf fldStatus = ParsingStatus.psErrorNonintegralIndex Then
		Parse = ParsingStatus.psErrorNonintegralIndex
		
	' ...or a successful parsing.
	Else
		Parse = ParsingStatus.psSuccess
	End If
	
	Exit Function
	
	
' Report a generic syntax error.
STX_ERROR:
	Parse = ParsingStatus.psError
End Function



' #############
' ## Support ##
' #############

' Close a field and record its elemental information.
Private Function EndField( _
	ByRef format As String, _
	ByRef e As ParsingElement, _
	ByRef cxt As ParsingContext, _
	ByRef nQuo As Long, _
	ByRef idxEsc As Boolean, _
	ByRef idxStart As Long, _
	ByRef idxStop As Long, _
	ByRef fmtStart As Long, _
	ByRef fmtStop As Long _
) As ParsingStatus
	Dim idxQuo As Boolean: idxQuo = False
	
	' Record the index.
	If e.HasIndex And idxStart < idxStop Then
		idxStop = idxStop - 1
		Dim idxLen As Long: idxLen = idxStop - idxStart + 1
		e.RawIndex = VBA.Mid(format, idxStart, idxLen)
		idxQuo = (nQuo = 1)
	End If
	
	' Record the format.
	If e.HasFormat And fmtStart < fmtStop Then
		fmtStart = fmtStart + 1
		Dim fmtLen As Long: fmtLen = fmtStop - fmtStart + 1
		e.Format = VBA.Mid(format, fmtStart, fmtLen)
	End If
	
	' Ignore a missing index.
	If Not e.HasIndex Then
		EndField = ParsingStatus.psSuccess
		GoTo RESET_VARS
		
	' Test for a key...
	ElseIf idxQuo Or idxEsc Then
		e.IndexIsKey = True
		GoTo RESET_VARS
		
	' ...or an integral index.
	Else
		On Error GoTo IDX_ERROR
		VBA.CLng e.Index
		
		On Error GoTo 0
		EndField = ParsingStatus.psSuccess
		GoTo RESET_VARS
		
IDX_ERROR:
		On Error GoTo 0
		EndField = ParsingStatus.psErrorNonintegralIndex
		GoTo RESET_VARS
	End If
	
	
' Reset the trackers.
RESET_VARS:
	cxt = ParsingContext.[_Unknown]
	' isQuo = False
	' isEsc = False
	
	nQuo = 0
	idxEsc = False
	idxStart = 0: idxStop = 0
	fmtStart = 0: fmtStop = 0
End Function
