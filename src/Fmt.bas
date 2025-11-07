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
	psErrorUnenclosedField = 1002	' ...or an incomplete field...
	psErrorUnenclosedQuote = 1003	' ...or an incomplete quote...
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
	Exists As Boolean	' Whether this index exists in its field.
	Syntax As String	' The syntax that was parsed to define this index.
	Start As Long		' Where that syntax begins in the original string...
	Stop As Long		' ...and where it ends.
	
	' The type of index:
	Kind As IndexKind
	Position As Long	' A positional integer...
	Key As String		' ...or a textual key.
End Type


' ...and the custom format...
Public Type peFieldFormat
	Exists As Boolean	' Whether this format exists in its field.
	Syntax As String	' The syntax that was parsed to define this format.
	Start As Long		' Where that syntax begins in the original string...
	Stop As Long		' ...and where it ends.
End Type


' ...of a field embedded in formatting.
Public Type peField
	Index As peFieldIndex	' Any index for this field...
	Format As peFieldFormat	' ...along with any format.
End Type


' Element for parsing plain text in formatting.
Public Type pePlain
	Text As String		' The text to display literally.
End Type


' Elements into which formats are parsed.
Public Type ParsingElement
	Syntax As String	' The syntax that was parsed to define this element.
	Start As Long		' Where that syntax begins in the original string...
	Stop As Long		' ...and where it ends.
	
	' The subtype which extends this element:
	Kind As ElementKind
	Plain as pePlain	' Plain text which displays literally...
	Field As peField	' ...or a field which embeds a value.
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
	Dim fmtLen As Long: fmtLen = VBA.Len(format)
	
	' Short-circuit for unformatted input.
	If fmtLen = 0 Then
		charIndex = 0
		Erase elements
		Parse = ParsingStatus.psSuccess
		Exit Function
	End If
	
	
	' Size to accommodate all (possible) elements.
	Dim eLen As Long: eLen = VBA.Int(fmtLen / 2) + 1
	Dim eUp As Long: eUp = base + eLen - 1
	ReDim elements(base To eUp)
	
	
	' Track the current context for parsing...
	Dim cxt As ParsingContext: cxt = ParsingContext.[_Unknown]
	Dim dfu As ParsingDefusal: dfu = ParsingDefusal.[_Off]
	
	' ...and the current depth of nesting...
	Dim fldDepth As Long: fldDepth = 0
	
	' ...and the current element...
	Dim eIdx As Long: eIdx = base - 1
	Dim eLen As Long: eLen = 0
	
	' ...and the current characters.
	Dim char As String
	Dim nQuo As Long: nQuo = 0
	Dim idxEsc As Boolean: idxEsc = False
	Dim idxStart As Long, idxStop As Long
	Dim fmtStart As Long, fmtStop As Long
	Dim fldStart As Long, fldStop As Long
	Dim endStatus As ParsingStatus: endStatus = ParsingStatus.psSuccess
	
	
	
	' #############
	' ## Parsing ##
	' #############
	
	' Catch generic syntax errors.
	On Error GoTo STX_ERROR
	
	' Scan and parse the format.
	For charIndex = 1 To fmtLen
		
		' Extract the current character.
		char = VBA.Mid$(format, charIndex, 1)
		
	' Revisit the character.
	SAME_CHAR:
		' Interpret this character in context.
		Select Case cxt
		
		
		
		' ##############
		' ## Inactive ##
		' ##############
		
		Case ParsingContext.[_Unknown]
			Select Case char
			
			' Parse into a field...
			Case openField
				' ...
				
			' ...or interpret as text.
			Case Else
				' ...
			End Select
			
			
			
		' ################
		' ## Plain Text ##
		' ################
		
		Case ParsingContext.pcPlain
			Select Case dfu
			
			' Quote "inert" text...
			Case ParsingDefusal.pdQuote
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					' ...
					
				' ...or continue quoting.
				Case Else
					' ...
				End Select
				
			' ...or escape literal text...
			Case ParsingDefusal.pdEscape
				' ...
				
			' ...or parse "active" text.
			Case Else
				Select Case char
				
				' Quote the next characters...
				Case openQuote
					' ...
					
				' ...escape the next character...
				Case escape
					' ...
					
				' ...or parse into a field...
				Case openField
					' ...
					
				' ...or display literally.
				Case Else
					' ...
				End Select
			End Select
			
			
			
		' ###########
		' ## Field ##
		' ###########
		
		Case ParsingContext.pcField
			Select Case char
			
			' Parse out of the field...
			Case closeField
				' ...
				
			' ...or parse into the format...
			Case separator
				' ...
				
			' ...or parse the index.
			Case Else
				' ...
			End Select
			
			
			
		' ###################
		' ## Field | Index ##
		' ###################
		
		Case ParsingContext.pcFieldIndex
			Select Case dfu
			
			' Quote "inert" symbol...
			Case ParsingDefusal.pdQuote
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					' ...
					
				' ...or continue quoting.
				Case Else
					' ...
				End Select
				
			' ...or escape literal symbol...
			Case ParsingDefusal.pdEscape
				' ...
				
			' ...or parse "active" symbol.
			Case Else
				Select Case char
				
				' Escape the next character...
				Case escape
					' ...
					
				' ...or nest into the field...
				Case openField
					' ...
					
				' ...or unnest out of the field...
				Case closeField
					' ...
					
				' ...or parse into a quoted key...
				Case openQuote
					' ...
					
				' ' ...or parse into a format...
				' Case separator
				' 	' ...
					
				' ...or display literally.
				Case Else
					' ...
				End Select
			End Select
			
			
			
		' ####################
		' ## Field | Format ##
		' ####################
		
		Case ParsingContext.pcFieldFormat
			Select Case dfu
			
			' Include quoted symbol...
			Case ParsingDefusal.pdQuote
				' ...
				
			' ...or include escaped symbol...
			Case ParsingDefusal.pdEscape
				' ...
				
			' ...but parse "active" symbol.
			Else
				Select Case char
				
				' Escape the next character...
				Case escape
					' ...
					
				' ...or nest into the field...
				Case openField
					' ...
					
				' ...or unnest out of the field...
				Case closeField
					' ...
					
				' ...or parse into a quoted key.
				Case openQuote
					' ...
				End Select
			End Select
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
		endStatus = EndField( _
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
		If endStatus = ParsingStatus.psErrorNonintegralIndex Then Exit Do
		endStatus = ParsingStatus.psSuccess
		
		' Increment the element.
		eIdx = eIdx + 1
		
		GoTo NEXT_CHAR
		
	' Increment the character.
	NEXT_CHAR:
		
	Next charIndex
	
	
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
		endStatus = EndField( _
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
	
	
	
	Select Case dfu
	
	' Report status: hanging escape...
	Case ParsingDefusal.pdEscape
		Parse = ParsingStatus.psErrorHangingEscape
		
	' ...or an unclosed quote...
	Case ParsingDefusal.pdQuote
		Parse = ParsingStatus.psErrorUnenclosedQuote
		
	Case Else
		' ...or an unclosed field...
		If fldDepth <> 0 Then
			Parse = ParsingStatus.psErrorUnenclosedField
			
		' ...or a index of the wrong type...
		ElseIf endStatus = ParsingStatus.psErrorNonintegralIndex Then
			Parse = ParsingStatus.psErrorNonintegralIndex
			
		' ...or a successful parsing.
		Else
			Parse = ParsingStatus.psSuccess
		End If
	End Select
	
	Exit Function
	
	
' Report a generic syntax error.
STX_ERROR:
	Parse = ParsingStatus.psError
End Function



' #############
' ## Support ##
' #############

' #######################
' ## Support | Parsing ##
' #######################

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
	' dfu = ParsingDefusal.[_Off]
	
	nQuo = 0
	idxEsc = False
	idxStart = 0: idxStop = 0
	fmtStart = 0: fmtStop = 0
End Function



' ########################
' ## Support | Elements ##
' ########################

' Reset an element.
Private Sub ParsingElement_Reset(ByRef pe As ParsingElement)
	Dim reset As ParsingElement
	Let pe = reset
End Sub


' Copy one element into another.
Private Sub ParsingElement_Copy(ByRef pe1 As ParsingElement, ByRef pe2 As ParsingElement)
	Let pe2.Syntax			= pe1.Syntax
	Let pe2.Start			= pe1.Start
	Let pe2.Stop			= pe1.Stop
	Let pe2.Kind			= pe1.Kind
	'      .Plain			     .Plain
	Let pe2.Plain.Text		= pe1.Plain.Text
	'      .Field			     .Field
	'      .Field.Index		     .Field.Index
	Let pe2.Field.Index.Exists	= pe1.Field.Index.Exists
	Let pe2.Field.Index.Syntax	= pe1.Field.Index.Syntax
	Let pe2.Field.Index.Start	= pe1.Field.Index.Start
	Let pe2.Field.Index.Stop	= pe1.Field.Index.Stop
	Let pe2.Field.Index.Kind	= pe1.Field.Index.Kind
	Let pe2.Field.Index.Position	= pe1.Field.Index.Position
	Let pe2.Field.Index.Key		= pe1.Field.Index.Key
	'      .Field.Format		     .Field.Format
	Let pe2.Field.Format.Exists	= pe1.Field.Format.Exists
	Let pe2.Field.Format.Syntax	= pe1.Field.Format.Syntax
	Let pe2.Field.Format.Start	= pe1.Field.Format.Start
	Let pe2.Field.Format.Stop	= pe1.Field.Format.Stop
End Sub
