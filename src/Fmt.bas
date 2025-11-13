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
	fmtVbFormat	' The Format() function in VBA.
	fmtXlText	' The Text() function in Excel.
End Enum


' Outcomes of parsing.
Public Enum ParsingStatus
	stsSuccess = 0			' Report success.
	stsError = 1000			' Report a general syntax error.
	stsErrorHangingEscape = 1001	' Report a hanging escape...
	stsErrorUnenclosedField = 1002	' ...or an incomplete field...
	stsErrorUnenclosedQuote = 1003	' ...or an incomplete quote...
	stsErrorNonintegralIndex = 1004	' ...or an index that is not an integer.
End Enum


' Kinds of elements which may be parsed.
Public Enum ElementKind
	[_Unknown]	' Uninitialized.
	elmPlain	' Plain text which is displayed as is.
	elmField	' Field that is formatted and embedded.
End Enum


' Ways to defuse literal symbols rather than interpreting them.
Private Enum ParsingDefusal
	[_Off]		' No defusal.
	dfuEscape	' Defuse only the next character...
	dfuQuote	' ...or all characters within quotes.
End Enum


' Kinds of indices for extracting values.
Private Enum IndexKind
	[_Unknown]	' Uninitialized.
	idxPosition	' Integer for a position...
	idxKey		' ...or text for a key.
End Enum


' Positional arguments passed to an embedded field.
Private Enum FieldArgument
	[_None]
	argIndex	' The index at which to extract the value.
	argFormat	' The formatting applied to the value.
	[_All]
End Enum



' ###########
' ## Types ##
' ###########

' An expression for parsing.
Public Type ParserExpression
	Syntax As String	' The syntax that was parsed to define this expression.
	Start As Long		' Where that syntax begins in the original code...
	Stop AS Long		' ...and where it ends.
End Type


' Element for parsing the index...
Public Type ParserIndex
	Exists As Boolean	' Whether this index exists in its field.
	Expression As ParsingExpression	' The expression that defines this index.
	
	' The type of index:
	Kind As IndexKind
	Position As Long	' A positional integer...
	Key As String		' ...or a textual key.
End Type


' ...and the custom format...
Public Type ParserFormat
	Exists As Boolean	' Whether this format exists in its field.
	Expression As ParsingExpression	' The expression that defines this format.
End Type


' ...of a field embedded in formatting.
Public Type ParserField
	Index As ParserIndex	' Any index for this field...
	Format As ParserFormat	' ...along with any format.
End Type


' Elements into which formats are parsed.
Public Type ParserElement
	Expression As ParsingExpression	' The expression that defines this element.
	
	' The subtype which extends this element:
	Kind As ElementKind
	Plain as String		' Plain text which displays literally...
	Field As ParserField	' ...or a field which embeds a value.
End Type



' #########
' ## API ##
' #########

' .
Public Function Parse( _
	ByRef format As String, _
	ByRef elements() As ParserElement, _
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
		Parse = ParsingStatus.stsSuccess
		Exit Function
	End If
	
	
	' Size to accommodate all (possible) elements.
	Dim eLen As Long: eLen = VBA.Int(fmtLen / 2) + 1
	Dim eUp As Long: eUp = base + eLen - 1
	ReDim elements(base To eUp)
	
	
	' Track the current context for parsing...
	Dim dfu As ParsingDefusal: dfu = ParsingDefusal.[_Off]
	
	' ...and the current depth of nesting...
	Dim fldDepth As Long: fldDepth = 0
	
	' ...and the current element...
	Dim eIdx As Long: eIdx = base - 1
	Dim e As ParserElement
	
	' ...and the current characters.
	Dim char As String
	Dim nQuo As Long: nQuo = 0
	Dim idxEsc As Boolean: idxEsc = False
	Dim endStatus As ParsingStatus: endStatus = ParsingStatus.stsSuccess
	
	
	
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
		Select Case e.Kind
		
		
		
		' ##############
		' ## Inactive ##
		' ##############
		
		Case ElementKind.[_Unknown]
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
		
		Case ElementKind.elmPlain
			Select Case dfu
			
			' Quote "inert" text...
			Case ParsingDefusal.dfuQuote
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					' ...
					
				' ...or continue quoting.
				Case Else
					' ...
				End Select
				
			' ...or escape literal text...
				' ...
			Case ParsingDefusal.dfuEscape
				
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
		
			Select Case char
		Case ElementKind.elmField
			
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
		End Select
		
		
		
	' #############
	' ## Control ##
	' #############
	
	' Parse out of the field.
	END_FIELD:
		' Record the elemental information...
		endStatus = Fld_Close(e.Field, _
			format := format, _
			nQuo := nQuo, _
			idxEsc := idxEsc _
		)
		
		' ...and short-circuit for an index of the wrong type.
		If endStatus = ParsingStatus.stsErrorNonintegralIndex Then Exit Do
		endStatus = ParsingStatus.stsSuccess
		
		' Increment the element.
		eIdx = eIdx + 1
		
		GoTo NEXT_CHAR
		
	' Increment the character.
	NEXT_CHAR:
		
	Next charIndex
	
	
' Escape the loop.
EXIT_LOOP:
	
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
	Select Case e.Kind
	Case ElementKind.elmField
		endStatus = Fld_Close(e.Field, _
			format := format, _
			nQuo := nQuo, _
			idxEsc := idxEsc _
		)
	End Select
	
	
	
	Select Case dfu
	
	' Report status: hanging escape...
	Case ParsingDefusal.dfuEscape
		Parse = ParsingStatus.stsErrorHangingEscape
		
	' ...or an unclosed quote...
	Case ParsingDefusal.dfuQuote
		Parse = ParsingStatus.stsErrorUnenclosedQuote
		
	Case Else
		' ...or an unclosed field...
		If fldDepth <> 0 Then
			Parse = ParsingStatus.stsErrorUnenclosedField
			
		' ...or a index of the wrong type...
		ElseIf endStatus = ParsingStatus.stsErrorNonintegralIndex Then
			Parse = ParsingStatus.stsErrorNonintegralIndex
			
		' ...or a successful parsing.
		Else
			Parse = ParsingStatus.stsSuccess
		End If
	End Select
	
	Exit Function
	
	
' Report a generic syntax error.
STX_ERROR:
	Parse = ParsingStatus.stsError
End Function



' #############
' ## Support ##
' #############

' #######################
' ## Support | Parsing ##
' #######################

' ' Reset any global trackers.
' Private Sub Reset( _
' 	Optional ByRef dfu As ParsingDefusal, _
' 	Optional ByRef fldDepth As Long, _
' 	Optional ByRef eIdx As Long, _
' 	Optional ByRef e As ParserElement, _
' 	Optional ByRef char As String, _
' 	Optional ByRef nQuo As Long, _
' 	Optional ByRef idxEsc As Boolean, _
' 	Optional ByRef endStatus As ParsingStatus _
' )
' 	dfu = ParsingDefusal.[_Off]
' 	fldDepth = 0
' 	eIdx = 0
' 	Elm_Reset e
' 	char = VBA.vbNullString
' 	nQuo = 0
' 	idxEsc = False
' 	endStatus = ParsingStatus.stsSuccess
' End Sub


' ' Save an element.
' Private Function Save( _
' 	ByRef format As String, _
' 	ByRef elements As ParserElement(), _
' 	ByRef eIdx As Long, _
' 	ByRef e As ParserElement, _
' 	ByRef nQuo As Long, _
' 	ByRef idxEsc As Boolean _
' ) As ParsingStatus
' 	Save = Elm_Close(e, format := format, nQuo := nQuo, idxEsc := idxEsc)
' 	Elm_Clone e, elements(eIdx)
' End Function


' ' Close an element and record its information.
' Private Sub Elm_Close(ByRef elm As ParserElement, _
' 	ByRef format As String _
' )
' 	' Record the syntax...
' 	If elm.Start <= elm.Stop Then
' 		Dim elmLen As Long: elmLen = elm.Stop - elm.Start + 1
' 		elm.Syntax = VBA.Mid$(format, elm.Start, elmLen)
' 		
' 	' ...or clear invalid information.
' 	Else
' 		elm.Start = 0
' 		elm.Stop = 0
' 	End If
' End Sub


' Close an element and record its information.
Private Function Elm_Close(ByRef elm As ParserElement, _
	ByRef format As String, _
	ByRef nQuo As Long, _
	ByRef idxEsc As Boolean _
) As ParsingStatus
	Dim status As ParsingStatus
	Elm_Close = ParsingStatus.stsSuccess
	
	' Record the syntax...
	If elm.Start <= elm.Stop Then
		Dim elmLen As Long: elmLen = elm.Stop - elm.Start + 1
		elm.Syntax = VBA.Mid$(format, elm.Start, elmLen)
		
	' ...or clear invalid information.
	Else
		elm.Start = 0
		elm.Stop = 0
	End If
	
	' Record any error when closing its extended (sub)element.
	Select Case elm.Kind
	Case ElementKind.elmField
		status = Fld_Close(elm.Field, format := format, nQuo := nQuo, idxEsc := idxEsc)
	Case Else
		status = ParsingStatus.stsSuccess
	End Select
	
	If Elm_Close = ParsingStatus.stsSuccess Then Elm_Close = status
End Function


' Close a field (sub)element and record its information...
Private Function Fld_Close(ByRef fld As ParserField, _
	ByRef format As String, _
	ByRef nQuo As Long, _
	ByRef idxEsc As Boolean _
) As ParsingStatus
	Dim status As ParsingStatus
	Fld_Close = ParsingStatus.stsSuccess
	
	' Record any error when closing its index...
	status = Idx_Close(fld.Index, format := format, nQuo := nQuo, idxEsc := idxEsc)
	If Fld_Close = ParsingStatus.stsSuccess Then Fld_Close = status
	
	' ...and its format.
	status = Fmt_Close(fld.Format, format := format)
	If Fld_Close = ParsingStatus.stsSuccess Then Fld_Close = status
End Function


' ...along with its index (sub)element...
Private Function Idx_Close(ByRef idx As ParserIndex, _
	ByRef format As String, _
	ByRef nQuo As Long, _
	ByRef idxEsc As Boolean _
) As ParsingStatus
	Dim idxQuo As Boolean: idxQuo = False
	
	' Record the index...
	If idx.Exists And idx.Start <= idx.Stop Then
		idx.Stop = idx.Stop - 1
		Dim idxLen As Long: idxLen = idx.Stop - idx.Start + 1
		idx.Syntax = VBA.Mid$(format, idx.Start, idxLen)
		idxQuo = (nQuo = 1)
		
	' ...or clear invalid information.
	Else
		idx.Start = 0
		idx.Stop = 0
	End If
	
	' Ignore a missing index.
	If Not idx.Exists Then
		Fld_Close = ParsingStatus.stsSuccess
		Exit Function
		
	' Test for a key...
	ElseIf idxQuo Or idxEsc Then
		idx.Kind = IndexKind.idxKey
		
		Fld_Close = ParsingStatus.stsSuccess
		Exit Function
		
	' ...or an integral index.
	Else
		On Error GoTo IDX_ERROR
		idx.Position = VBA.CLng(idx.Key)
		On Error GoTo 0
		
		idx.Kind = IndexKind.idxPosition
		idx.Key = VBA.vbNullString
		
		Fld_Close = ParsingStatus.stsSuccess
		Exit Function
		
IDX_ERROR:
		On Error GoTo 0
		' idx.Kind = IndexKind.[_Unknown]
		
		Fld_Close = ParsingStatus.stsErrorNonintegralIndex
		Exit Function
	End If
End Function


' ...and its format (sub)element.
Private Function Fmt_Close(ByRef fmt As ParserFormat, _
	ByRef format As String _
) As ParsingStatus
	' Record the format...
	If fmt.Exists And fmt.Start <= fmt.Stop Then
		fmt.Start = fmt.Start + 1
		Dim fmtLen As Long: fmtLen = fmt.Stop - fmt.Start + 1
		fmt.Syntax = VBA.Mid$(format, fmt.Start, fmtLen)
		
	' ...or clear invalid information.
	Else
		fmt.Start = 0
		fmt.Stop = 0
	End If
	
	' This should always work.
	Fmt_Close = ParsingStatus.stsSuccess
End Function



' ########################
' ## Support | Elements ##
' ########################

' Reset an element.
Private Sub Elm_Reset(ByRef elm As ParserElement)
	Dim reset As ParserElement
	Let elm = reset
End Sub


' Clone one expression into another.
Private Sub Expr_Clone(ByRef expr1 As ParserExpression, ByRef expr2 As ParserExpression)
	Let expr2.Syntax = expr1.Syntax
	Let expr2.Start  = expr1.Start
	Let expr2.Stop   = expr1.Stop
End Sub


' Clone one element into another.
Private Sub Elm_Clone(ByRef elm1 As ParserElement, ByRef elm2 As ParserElement)
	Expr_Clone elm1.Expression, elm2.Expression
	
	Let elm2.Kind   = elm1.Kind
	Let elm2.Plain	= elm1.Plain
	Fld_Clone elm1.Field, elm2.Field
End Sub


' Clone one field (sub)element into another...
Private Sub Fld_Clone(ByRef fld1 As ParserField, ByRef fld2 As ParserField)
	Idx_Clone fld1.Index,  fld2.Index
	Fmt_Clone fld1.Format, fld2.Format
End Sub


' ...and its index (sub)element into another...
Private Sub Idx_Clone(ByRef idx1 As ParserIndex, ByRef idx2 As ParserIndex)
	Let idx2.Exists   = idx1.Exists
	Expr_Clone idx1.Expression, idx2.Expression
	Let idx2.Kind     = idx1.Kind
	Let idx2.Position = idx1.Position
	Let idx2.Key      = idx1.Key
End Sub


' ...and its format (sub)element into another.
Private Sub Fmt_Clone(ByRef fmt1 As ParserFormat, ByRef fmt2 As ParserFormat)
	Let fmt2.Exists = fmt1.Exists
	Expr_Clone fmt1.Expression, fmt2.Expression
End Sub
