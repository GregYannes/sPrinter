Attribute VB_Name = "sPrinter"



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

Public Const MOD_NAME As String = "sPrinter"

Public Const MOD_VERSION As String = ""

Public Const MOD_REPO As String = "https://github.com/GregYannes/sPrinter"



' ###############
' ## Constants ##
' ###############

' Symbols for parsing syntax.
Private Const SYM_ESC As String = "\"			' Escape the next character.
Private Const SYM_FLD_OPEN As String = "{"		' Embed a field for formatting...
Private Const SYM_FLD_CLOSE As String = "}"		' ...and enclose that field.
Private Const SYM_QUO_OPEN As String = """"		' Quote the next several characters...
Private Const SYM_QUO_CLOSE As String = SYM_QUO_OPEN	' ...and enclose that quote.
Private Const SYM_SEP As String = ":"			' Separate the arguments in a field.



' ##################
' ## Enumerations ##
' ##################

' Engine used for formatting.
Public Enum FormatMode
	[_Unknown] = 0	' Uninitialized.
	fmtVbFormat	' The Format() function in VBA.
	fmtXlText	' The Text() function in Excel.
End Enum


' Context from which Message() is called.
Private Enum CallingContext
	[_Unknown] = 0	' Uninitialized.
	cxtVBA		' The VBA environment.
	cxtExcel	' An Excel worksheet.
End Enum


' ' Syntax for parsing.
' Private Enum ParsingSymbol
' '	=============	==========================	  =============		=========	====================================
' '	Label		Code				  Name			Character	Description
' '	=============	==========================	  =============		=========	====================================
' 	symEscape     =	                        92	' Backslash		\		Escape the next character.
' 	symOpenField  =	                       173	' Opening brace		{		Embed a field for formatting...
' 	symCloseField =	                       175	' Closing brace		}		...and enclose that field.
' 	symOpenQuote  =	                        34	' Double quotes		"		Quote the next several characters...
' 	symCloseQuote =	ParsingSymbol.symOpenQuote	' Double quotes		"		...and enclose that quote.
' 	symSeparator  =	                        58	' Colon			:		Separate the arguments in a field.
' End Enum


' Outcomes of parsing.
Public Enum ParsingStatus
	stsSuccess                =    0	' Report success.
	stsError                  = 1000	' Report a general syntax error.
	stsErrorHangingEscape     = 1001	' Report a hanging escape...
	stsErrorUnenclosedQuote   = 1002	' ...or an incomplete quote...
	stsErrorImbalancedNesting = 1003	' ...or an imbalanced nesting...
	stsErrorInvalidIndex      = 1004	' ...or an index that is not an integer.
' 	stsErrorNonexistentIndex  = 1005	' Report an index that does not exist in the data...
' 	stsErrorInvalidFormat     = 1006	' ...or a format that is invalid for Format() or Text().
End Enum


' Kinds of elements which may be parsed.
Public Enum ElementKind
	[_Unknown] = 0	' Uninitialized.
	elmPlain	' Plain text which is displayed as is.
	elmField	' Field that is formatted and embedded.
End Enum


' Ways to defuse literal symbols rather than interpreting them.
' NOTE: These may be combined (+) so they apply simultaneously.
Private Enum ParsingDefusal
	[_Off]    = 0		' No defusal.
	dfuEscape = 2 ^ 0	' Defuse only the next character...
	dfuQuote  = 2 ^ 1	' ...or all characters within quotes...
	dfuNest   = 2 ^ 2	' ...or all expressions within a nested field.
End Enum


' Positional arguments passed to an embedded field.
Private Enum FieldArgument
	[_None]					' No arguments.
	argIndex				' The index at which to extract the value.
	argFormat				' The formatting applied to the value.
	[_All]					' All arguments.
	
	[_First] = FieldArgument.[_None] + 1	' The first argument.
	[_Last]  = FieldArgument.[_All] - 1	' The last argument.
End Enum


' Ways to interpret (negative) positional indices.
Public Enum PositionKind
	[_Unknown] = 0	' Uninitialized.
	posAbsolute	' Negative index (-1) is extracted directly...
	posRelative	' ...or measured (1st) from the end.
End Enum



' ###########
' ## Types ##
' ###########

' An expression for parsing.
Public Type ParserExpression
	Syntax As String	' The syntax that was parsed to define this expression.
	Start As Long		' Where that syntax begins in the original code...
	Stop As Long		' ...and where it ends.
End Type


' Element for parsing a field embedded in formatting.
Public Type ParserField
	Index As Variant		' The index to extract the value.
	Format As String		' The formatting code applied to the value.
End Type


' Elements into which formats are parsed.
Public Type ParserElement
	Kind As ElementKind		' The subtype which extends this element:
	Plain As String			'   - Plain text which displays literally...
	Field As ParserField		'   - ...or a field which embeds a value.
End Type



' #########
' ## API ##
' #########

' ######################
' ## API | Formatting ##
' ######################

' Display a value with a certain formatting code and engine.
Public Function Format2( _
	ByRef value As Variant, _
	Optional ByRef format As String, _
	Optional ByVal mode As FormatMode = FormatMode.[_Unknown], _
	Optional ByVal firstDayOfWeek As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbSunday, _
	Optional ByVal firstWeekOfYear As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbFirstJan1 _
) As String
	' Default to inferring the mode from calling context.
	If mode = FormatMode.[_Unknown] Then
		Dim cxt As CallingContext: cxt = Context()
		Select Case cxt
		
		' Calling from VBA will default to its native Format()...
		Case CallingContext.cxtVBA:   mode = FormatMode.fmtVbFormat
		
		' ...while calling from Excel will default to its native TEXT()...
		Case CallingContext.cxtExcel: mode = FormatMode.fmtXlText
		
		' ...but the more rigorous Format() is our failsafe.
		Case Else:                    mode = FormatMode.fmtVbFormat
		End Select
	End If
	
	
	' Format via the desired mode.
	Select Case mode
	Case FormatMode.fmtVbFormat
		Format2 = VBA.Format( _
			Expression := value, _
			Format := format, _
			FirstDayOfWeek := firstDayOfWeek, _
			FirstWeekOfYear := firstWeekOfYear _
		)
		
	Case FormatMode.fmtXlText
		Format2 = Application.WorksheetFunction.Text( _
			Arg1 := value, _
			Arg2 := format _
		)
	End Select
End Function


' Embed values (with formatting) within a message...
Public Function Message( _
	ByRef format As String, _
	Optional ByRef data As Variant, _
	Optional ByVal mode As FormatMode = FormatMode.[_Unknown], _
	Optional ByVal firstDayOfWeek As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbSunday, _
	Optional ByVal firstWeekOfYear As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbFirstJan1, _
	Optional ByVal position As PositionKind = PositionKind.posAbsolute, _
	Optional ByVal escape As Variant = SYM_ESC, _
	Optional ByVal openField As String = SYM_FLD_OPEN, _
	Optional ByVal closeField As String = SYM_FLD_CLOSE, _
	Optional ByVal openQuote As String = SYM_QUO_OPEN, _
	Optional ByVal closeQuote As String = SYM_QUO_CLOSE, _
	Optional ByVal separator As String = SYM_SEP _
) As String
	' Short-circuit for blank format.
	If format = VBA.vbNullString Then
		Message = VBA.vbNullString
		Exit Function
	End If
	
	' ...
End Function


' ...and print that message to the console.
Public Function Print2( _
	ByRef format As String, _
	Optional ByRef data As Variant, _
	Optional ByVal mode As FormatMode = FormatMode.[_Unknown], _
	Optional ByVal firstDayOfWeek As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbSunday, _
	Optional ByVal firstWeekOfYear As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbFirstJan1, _
	Optional ByVal position As PositionKind = PositionKind.posAbsolute, _
	Optional ByVal escape As Variant = SYM_ESC, _
	Optional ByVal openField As String = SYM_FLD_OPEN, _
	Optional ByVal closeField As String = SYM_FLD_CLOSE, _
	Optional ByVal openQuote As String = SYM_QUO_OPEN, _
	Optional ByVal closeQuote As String = SYM_QUO_CLOSE, _
	Optional ByVal separator As String = SYM_SEP _
) As String
	Print2 = Message( _
		format := format, _
		data := data, _
		mode := mode, _
		firstDayOfWeek := firstDayOfWeek, _
		firstWeekOfYear := firstWeekOfYear, _
		position := position, _
		escape := escape, _
		openField := openField, _
		closeField := closeField, _
		openQuote := openQuote, _
		closeQuote := closeQuote, _
		separator := separator _
	)
	
	Debug.Print Print2
End Function



' ' ####################
' ' ## API | Printing ##
' ' ####################
' 
' ' .
' Public Function Printf() As String
' 	' ...
' End Function
' 
' 
' ' .
' Public Function vPrintf() As String
' 	' ...
' End Function
' 
' 
' ' .
' Public Function sPrintf() As String
' 	' ...
' End Function
' 
' 
' ' .
' Public Function vsPrintf() As String
' 	' ...
' End Function



' ###################
' ## API | Parsing ##
' ###################

' Parse a format string as an array of syntax elements.
Public Function Parse( _
	ByRef format As String, _
	Optional ByVal base As Long = 1, _
	Optional ByVal escape As Variant = SYM_ESC, _
	Optional ByVal openField As String = SYM_FLD_OPEN, _
	Optional ByVal closeField As String = SYM_FLD_CLOSE, _
	Optional ByVal openQuote As String = SYM_QUO_OPEN, _
	Optional ByVal closeQuote As String = SYM_QUO_CLOSE, _
	Optional ByVal separator As String = SYM_SEP _
) As ParserElement()
	' Validate the symbols for syntax.
	CheckSyms _
		escape := escape, _
		openField := openField, _
		closeField := closeField, _
		openQuote := openQuote, _
		closeQuote := closeQuote, _
		separator := separator
	
	' Parse the format and record the outcome.
	Dim expression As ParserExpression
	Dim status As ParsingStatus
	Parse0 _
		format := format, _
		base := base, _
		escape := escape, _
		openField := openField, _
		closeField := closeField, _
		openQuote := openQuote, _
		closeQuote := closeQuote, _
		separator := separator, _
		elements := Parse, _
		expression := expression, _
		status := status
	
	' Raise any error that occurred while parsing.
	If status <> ParsingStatus.stsSuccess Then
		Err_Parsing _
			status := status, _
			expression := expression, _
			escape := escape, _
			openField := openField, _
			closeField := closeField, _
			openQuote := openQuote, _
			closeQuote := closeQuote, _
			separator := separator
	End If
End Function



' #################
' ## Diagnostics ##
' #################

' Throw a parsing error with granular information.
Private Sub Err_Parsing( _
	ByVal status As ParsingStatus, _
	ByRef expression As ParserExpression, _
	ByVal escape As String, _
	ByVal openField As String, _
	ByVal closeField As String, _
	ByVal openQuote As String, _
	ByVal closeQuote As String, _
	ByVal separator As String _
)
	' Define the format for cardinal numbers: 1st, 2nd, 3rd, 4th, etc.
	Const ORD_FMT As String = "#,##0"
	
	' Define the horizontal ellipsis: "…"
	#If Mac Then
		Const ETC_SYM As Long = 201
	#Else
		Const ETC_SYM As Long = 133
	#End If
	
	Dim etc As String: etc = VBA.Chr(ETC_SYM)
	
	
	' Describe where the erroneous syntax occurs.
	Dim description As String, position As String
	Dim startPos As String, stopPos As String
	
	If expression.Start > 0 Then
		startPos = Num_Ordinal(expression.Start, format := ORD_FMT)
		stopPos = Num_Ordinal(expression.Stop, format := ORD_FMT)
		
		If expression.Start < expression.Stop Then
			position = "between the " & startPos & " and " & stopPos & " characters"
		ElseIf expression.Start = expression.Stop Then
			position = "at the " & startPos & " character"
		Else
			position = "following the " & stopPos & " character"
		End If
	End If
	
	' Generate a relevant description of the error.
	Select Case status
	Case ParsingStatus.stsError
		description = "An error occurred when parsing the message format"
		If position <> VBA.vbNullString Then description = description & ", " & position
		description = description & "."
		
	Case ParsingStatus.stsErrorHangingEscape
		description = "The message format contains a hanging escape (" & escape & ")"
		If position <> VBA.vbNullString Then description = description & " " & position
		description = description & "."
		
	Case ParsingStatus.stsErrorUnenclosedQuote
		description = "The message format contains an unenclosed quote (" & openQuote & etc & closeQuote & ")"
		If position <> VBA.vbNullString Then description = description & " " & position
		description = description & "."
		
	Case ParsingStatus.stsErrorImbalancedNesting
		description = "The message format contains an imbalanced field nesting (" & openField & etc & closeField & ")"
		If position <> VBA.vbNullString Then description = description & " " & position
		description = description & "."
		
	Case ParsingStatus.stsErrorInvalidIndex
		description = "The message format contains an invalid field index"
		If position <> VBA.vbNullString Then description = description & ", " & position
		description = description & ": " & expression.Syntax
		
	Case Else: Exit Sub
	End Select
	
	' Raise the error.
	Err.Raise _
		Number := status, _
		Description := description
End Sub


' ' Throw an error for a nonexistent index.
' Private Sub Err_Index(ByRef index As Variant)
' 	ByRef position As PositionKind _
' 	
' 	
' 	' Display the index in detail.
' 	Dim idxCode As String, idxKind As String  ' , posKind As String
' 	FormatIndex _
' 		idx := index, _
' 		idxCode := idxCode, _
' 		ord := False, _
' 		idxKind := idxKind, _
' 		pos := position  ' , _
' 		posKind := posKind
' 	
' 	' Generate a relevant description of the error.
' 	Dim description As String
' 	description = "This"
' 	' If posKind <> VBA.vbNullString Then description = description & " (" & posKind & ")"
' 	description = description & " " & idxKind & " does not exist in the data: " & idxCode
' 	
' 	' Raise the error.
' 	Err.Raise _
' 		Number := ParsingStatus.stsErrorNonexistentIndex, _
' 		Description := description
' End Sub


' ' Throw an error for an invalid format.
' Private Sub Err_Format( _
' 	ByRef index As Variant, _
' 	ByRef format As String, _
' 	Optional ByVal position As PositionKind _
' )
' 	' Display the index in detail.
' 	Dim idxCode As String, idxKind As String  ' , posKind As String
' 	FormatIndex _
' 		idx := index, _
' 		idxCode := idxCode, _
' 		ord := False, _
' 		idxKind := idxKind, _
' 		pos := position  ' , _
' 		posKind := posKind
' 	
' 	' Generate a relevant description of the error.
' 	Dim description As String
' 	description = "The value from this"
' 	' If posKind <> VBA.vbNullString Then description = description & " (" & posKind & ")"
' 	description = description & " " & idxKind & " (" & idxCode & ") cannot be displayed in this format: " & format
' 	
' 	' Raise the error.
' 	Err.Raise _
' 		Number := ParsingStatus.stsErrorInvalidFormat, _
' 		Description := description
' End Sub


' Throw an error for a blank parsing symbol.
Private Sub Err_BlankSym()
	' Detail the error.
	Const ERR_NUM As Long = 5
	Const ERR_DESC As String = "Whitespace may not be used as a formatting symbol."
	
	' Raise the error.
	Err.Raise _
		Number := ERR_NUM, _
		Description := ERR_DESC
End Sub


' Throw an error for duplicate parsing symbols.
Private Sub Err_DuplicateSyms( _
	ByVal sym As String, _
	ByVal openQuote As String, _
	ByVal closeQuote As String _
)
	' Define the error.
	Const ERR_NUM As Long = 5
	
	' Define the horizontal ellipsis: "…"
	#If Mac Then
		Const ETC_SYM As Long = 201
	#Else
		Const ETC_SYM As Long = 133
	#End If
	
	Dim etc As String: etc = VBA.Chr(ETC_SYM)
	
	
	' Generate a relevant description of the error.
	Dim description As String
	description = "The same formatting symbol (""" &  sym & """) may not be used twice"
	description = description & "; " & "aside from quotes (" & openQuote & etc & closeQuote & ") if you so specify"
	description = description & "."
	
	' Raise the error.
	Err.Raise _
		Number := ERR_NUM, _
		Description := description
End Sub



' ###############
' ## Utilities ##
' ###############

' Test if a combination (dfuNest + dfuEscape) includes a particular enumeration (dfuEscape).
Public Function Enum_Has(ByRef enum1 As Long, ByRef enum2 As Long) As Boolean
	Enum_Has = enum1 And enum2
End Function


' Display the ordinal (3rd) of an integer (3).
Public Function Num_Ordinal(ByVal num As Long, _
	Optional ByRef format As String _
) As String
	Const NUM_BASE As Integer = 10
	
	' Determine the proper suffix...
	Dim sfx As String
	Select Case (num Mod NUM_BASE)
	Case 1:		sfx = "st"
	Case 2:		sfx = "nd"
	Case 3:		sfx = "rd"
	Case Else:	sfx = "th"
	End Select
	
	' ...and the proper prefix.
	Dim pfx As String
	If format = VBA.vbNullString Then
		pfx = VBA.CStr(num)
	Else
		pfx = VBA.Format(num, Format := format)
	End If
	
	' Combine them and return the result.
	Num_Ordinal = pfx & sfx
End Function



' ######################
' ## Utilities | Text ##
' ######################

' Remove characters from the end(s) of a string.
Public Function Txt_Crop(ByVal txt As String, _
	Optional ByVal nLeft As Long = 0, _
	Optional ByVal nRight As Long = 0 _
)
	Dim n As Long
	
	If nRight > 0 Then
		' Record the initial length...
		n = VBA.Len(txt)
		
		' ...and truncate the suffix.
		nRight = Application.WorksheetFunction.Min(nRight, n)
		txt = VBA.Left$(txt, n - nRight)
	End If
	
	If nLeft > 0 Then
		' Record the remaining length...
		n = VBA.Len(txt)
		
		' ...and truncate the prefix...
		nLeft = Application.WorksheetFunction.Min(nLeft, n)
		txt = VBA.Right$(txt, n - nLeft)
	End If
	
	' Return the result.
	Txt_Crop = txt
End Function



' #############
' ## Support ##
' #############

' Determine the context from which the current function was called.
Private Function Context() As CallingContext
	If TypeOf Application.Caller Is Range Then
		Context = CallingContext.cxtExcel
	ElseIf VBA.IsError(Application.Caller) Then
		If VBA.CLng(Application.Caller) = Excel.XlCVError.xlErrRef Then
			Context = CallingContext.cxtVBA
		Else
			Context = CallingContext.[_Unknown]
		End If
	Else
		Context = CallingContext.[_Unknown]
	End If
End Function



' ##########################
' ## Support | Formatting ##
' ##########################

' ' Display the index for a field...
' Private Sub FormatIndex( _
' 	ByRef idx As Variant, _
' 	Optional ByRef idxCode As String, _
' 	Optional ByVal ord As Boolean = False, _
' 	Optional ByRef idxKind As String, _
' 	Optional ByVal pos As PositionKind, _
' 	Optional ByRef posKind As String _
' )
' 	' Define the format for numeric indices.
' 	Const IDX_FMT As String = "#,##0"
' 	
' 	' Define how a key is displayed.
' 	Dim KEY_OPEN As String = """"
' 	Dim KEY_CLOSE As String = """"
' 	
' 	
' 	Select Case VBA.VarType(idx)
' 	Case VBA.VbVarType.vbLong
' 		If ord Then
' 			idxCode = Num_Ordinal(idx, format := IDX_FMT)
' 		Else
' 			idxCode = VBA.Format(idx, Format := IDX_FMT)
' 		End If
' 		
' 		idxKind = "Position"
' 		
' 		Select Case pos
' 		Case PositionKind.posAbsolute: posKind = "Absolute"
' 		Case PositionKind.posRelative: posKind = "Relative"
' 		End Select
' 		
' 	Case VBA.VbVarType.vbString
' 		idxCode = KEY_OPEN & idx & KEY_CLOSE
' 		idxKind = "Key"
' 	End Select
' End Sub



' ##########################
' ## Support | Validation ##
' ##########################

' Validate input (text or code) for a parsing symbol.
Private Sub CheckSym(ByRef x As Variant)
	' Extract the first character from a string...
	If VBA.VarType(x) = VBA.VbVarType.vbString Then
		x = VBA.Left$(x, 1)
		
	' ...or convert a code into its character.
	Else
		#If Mac Then
			x = VBA.Chr(x)
		#Else
			x = VBA.ChrW(x)
		#End If
	End If
	
	' Ensure the symbol is not whitespace...
	x = Application.WorksheetFunction.Clean$(x)
	x = VBA.Trim$(x)
	If x = VBA.vbNullString Then GoTo BLANK_ERROR
	
	' ...and return the result.
	Exit Sub
	
	
' Throw an error for whitespace.
BLANK_ERROR:
	Err_BlankSym
End Sub


' Validate inputs (texts or codes) for all parsing symbols.
Private Sub CheckSyms( _
	ByRef escape As Variant, _
	ByRef openField As Variant, _
	ByRef closeField As Variant, _
	ByRef openQuote As Variant, _
	ByRef closeQuote As Variant, _
	ByRef separator As Variant _
)
	' Validate individual symbols.
	CheckSym escape
	CheckSym openField
	CheckSym closeField
	CheckSym openQuote
	CheckSym closeQuote
	CheckSym separator
	
	
	' Validate uniqueness across symbols...
	Dim syms As Collection: Set syms = New Collection
	Dim sym As String
	
	On Error GoTo DUP_ERROR
	sym = escape:     syms.Add False, key := sym
	sym = openField:  syms.Add False, key := sym
	sym = closeField: syms.Add False, key := sym
	sym = openQuote:  syms.Add False, key := sym
	
	' ...except between quotes.
	If openQuote <> closeQuote Then
		sym = closeQuote
		syms.Add False, key := closeQuote
	End If
	
	sym = separator:  syms.Add False, key := sym
	On Error GoTo 0
	
	' Conclude validation successfully.
	Exit Sub
	
	
' Report an error for clashing symbols.
DUP_ERROR:
	Err_DuplicateSyms _
		sym := sym, _
		openQuote := openQuote, _
		closeQuote := closeQuote
End Sub



' #######################
' ## Support | Parsing ##
' #######################

' Parse a format string (without guardrails) and record granular details.
Private Sub Parse0( _
	ByRef format As String, _
	ByRef base As Long, _
	ByVal escape As String, _
	ByVal openField As String, _
	ByVal closeField As String, _
	ByVal openQuote As String, _
	ByVal closeQuote As String, _
	ByVal separator As String, _
	ByRef elements() As ParserElement, _
	ByRef expression As ParserExpression, _
	Optional ByRef status As ParsingStatus _
)
	' ###########
	' ## Setup ##
	' ###########
	
	' Default to success.
	status = ParsingStatus.stsSuccess
	
	' Record the format length.
	Dim fmtLen As Long: fmtLen = VBA.Len(format)
	
	' Short-circuit for unformatted input.
	If fmtLen = 0 Then
		Erase elements
		Expr_Reset expression
		status = ParsingStatus.stsSuccess
		Exit Sub
	End If
	
	
	' Size to accommodate all (possible) elements.
	Dim eLen As Long: eLen = VBA.Int(fmtLen / 2) + 1
	Dim eUp As Long: eUp = base + eLen - 1
	ReDim elements(base To eUp)
	
	
	' Track the current context for parsing...
	Dim dfu As ParsingDefusal: dfu = ParsingDefusal.[_Off]
	Dim depth As Long: depth = 0
	
	' ...and the current element...
	Dim eIdx As Long: eIdx = base
	Dim e As ParserElement
	Expr_Reset expression
	
	' ...and the current (field) argument...
	Dim args(FieldArgument.[_First] To FieldArgument.[_Last]) As ParserExpression
	Dim argIdx As FieldArgument: argIdx = FieldArgument.[_None]
	Dim arg As ParserExpression
	Dim idxDfu As ParserExpression
	Dim idxEsc As Long: idxEsc = 0
	
	' ...and the current characters.
	Dim charIndex As Long
	Dim char As String
	
	
	
	' #############
	' ## Parsing ##
	' #############
	
	' Catch generic syntax errors.
	On Error GoTo STX_ERROR
	
	' Scan and parse the format.
	For charIndex = 1 To fmtLen
		
		' Extract the current character.
		char = VBA.Mid$(format, charIndex, 1)
		
	' Revisit the character from scratch.
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
				' Nest deeper into the field.
				depth = depth + 1
				
				' Identify the element as a field.
				e.Kind = ElementKind.elmField
				
				' Locate the element.
				expression.Start = charIndex
				expression.Stop = expression.Start
				
				' Advance to the next character in the field.
				GoTo NEXT_CHAR
				
			' ...or interpret as text.
			Case Else
				' Identify the element as plaintext.
				e.Kind = ElementKind.elmPlain
				
				' Locate the element.
				expression.Start = charIndex
				expression.Stop = expression.Start
				
				' Revisit this character in plaintext.
				expression.Stop = expression.Stop - 1
				GoTo SAME_CHAR
			End Select
			
			
			
		' ################
		' ## Plain Text ##
		' ################
		
		Case ElementKind.elmPlain
			' Escape a literal character...
			If Enum_Has(dfu, ParsingDefusal.dfuEscape) Then
				' Deactivate escaping.
				dfu = dfu - ParsingDefusal.dfuEscape
				
				' Extend the location...
				expression.Stop = expression.Stop + 1
				
				' ...and the contents.
				e.Plain = e.Plain & char
				
				' Advance to the next character.
				GoTo NEXT_CHAR
				
			' ...or quote "inert" text...
			ElseIf Enum_Has(dfu, ParsingDefusal.dfuQuote) Then
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					' Deactivate quoting.
					dfu = dfu - ParsingDefusal.dfuQuote
					
					' Extend the location.
					expression.Stop = expression.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or continue quoting.
				Case Else
					' Extend the location...
					expression.Stop = expression.Stop + 1
					
					' ...and the contents.
					e.Plain = e.Plain & char
					
					' Advance to the next character.
					GoTo NEXT_CHAR
				End Select
				
			' ...or parse "active" expressions.
			Else
				Select Case char
				
				' Escape the next character...
				Case escape
					' Activate escaping.
					dfu = dfu + ParsingDefusal.dfuEscape
					
					' Extend the location.
					expression.Stop = expression.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or quote the next characters...
				Case openQuote
					' Activate quoting.
					dfu = dfu + ParsingDefusal.dfuQuote
					
					' Extend the location.
					expression.Stop = expression.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or parse into a field...
				Case openField
					' Save (and reset) this (plaintext) element to the array.
					Elm_Clone e, elements(eIdx)
					Elm_Reset e
					
					' Advance to the next element.
					eIdx = eIdx + 1
					
					' Reset the global trackers.
					Expr_Reset expression
					
					' Revisit this character from scratch.
					GoTo SAME_CHAR
					
				' ...or display literally.
				Case Else
					' Extend the location...
					expression.Stop = expression.Stop + 1
					
					' ...and the contents.
					e.Plain = e.Plain & char
					
					' Advance to the next character.
					GoTo NEXT_CHAR
				End Select
			End If
			
			
			
		' ###########
		' ## Field ##
		' ###########
		
		Case ElementKind.elmField
			' Escape a literal character...
			If Enum_Has(dfu, ParsingDefusal.dfuEscape) Then
				' Deactivate escaping.
				dfu = dfu - ParsingDefusal.dfuEscape
				
				' Extend the location of this field...
				expression.Stop = expression.Stop + 1
				
				' ...and of this argument...
				arg.Stop = arg.Stop + 1
				
				' ...along with its (defused) contents.
				arg.Syntax = arg.Syntax & char
				
				' Advance to the next character.
				GoTo NEXT_CHAR
				
			' ...or quote "inert" text...
			ElseIf Enum_Has(dfu, ParsingDefusal.dfuQuote) Then
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					' Deactivate quoting.
					dfu = dfu - ParsingDefusal.dfuQuote
						
					' Locate any quoting in the index argument.
					If depth = 1 And argIndex = FieldArgument.argIndex Then idxDfu.Stop = charIndex
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument.
					arg.Stop = arg.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or continue quoting.
				Case Else
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument...
					arg.Stop = arg.Stop + 1
					
					' ...along with its (defused) contents.
					arg.Syntax = arg.Syntax & char
					
					' Advance to the next character.
					GoTo NEXT_CHAR
				End Select
				
			' ...or nest expressions...
			ElseIf Enum_Has(dfu, ParsingDefusal.dfuNest) Then
				Select case char
				
				' Escape the next character...
				Case escape
					' Activate escaping.
					dfu = dfu + ParsingDefusal.dfuEscape
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument.
					arg.Stop = arg.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or quote the next characters...
				Case openQuote
					' Activate quoting.
					dfu = dfu + ParsingDefusal.dfuQuote
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument.
					arg.Stop = arg.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or nest deeper...
				Case openField
					' Nest deeper into the nesting.
					depth = depth + 1
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument...
					arg.Stop = arg.Stop + 1
					
					' ...along with its (defused) contents.
					arg.Syntax = arg.Syntax & char
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or unnest shallower...
				Case closeField
					' Unnest shallower out of the nesting.
					depth = depth - 1
					
					' Extend any (defused) contents that remain nested...
					If depth > 1 Then
						arg.Syntax = arg.Syntax & char
						
					' ...but otherwise deactivate nesting...
					Else
						dfu = dfu - ParsingDefusal.dfuNest
						
						' ...and locate any nesting in the index argument.
						If argIdx = FieldArgument.argIndex Then idxDfu.Stop = charIndex
					End If
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument.
					arg.Stop = arg.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or display literally.
				Case Else
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument...
					arg.Stop = arg.Stop + 1
					
					' ...along with its (defused) contents.
					arg.Syntax = arg.Syntax & char
					
					' Advance to the next character.
					GoTo NEXT_CHAR
				End Select
				
			' ...or parse "active" expressions.
			Else
				Select Case char
				
				' Escape the next character...
				Case escape
					' Confirm the argument...
					If argIdx = FieldArgument.[_None] Then
						argIdx = argIdx + 1
						arg.Start = charIndex
						arg.Stop = arg.Start
						
					' ...or extend its location.
					Else
						arg.Stop = arg.Stop + 1
					End If
					
					' Activate escaping.
					dfu = dfu + ParsingDefusal.dfuEscape
					
					' Note any escaping in the index argument.
					If depth = 1 And argIdx = FieldArgument.argIndex And idxEsc = 0 Then idxEsc = charIndex
					
					' Extend the location of this field.
					expression.Stop = expression.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or quote the next characters...
				Case openQuote
					' Confirm the argument...
					If argIdx = FieldArgument.[_None] Then
						argIdx = argIdx + 1
						arg.Start = charIndex
						arg.Stop = arg.Start
						
					' ...or extend its location.
					Else
						arg.Stop = arg.Stop + 1
					End If
					
					' Activate quoting.
					dfu = dfu + ParsingDefusal.dfuQuote
					
					' Locate any quoting in the index argument.
					If depth = 1 And argIndex = FieldArgument.argIndex Then idxDfu.Start = charIndex
					
					' Extend the location of this field.
					expression.Stop = expression.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or nest the next syntax...
				Case openField
					' Confirm the argument...
					If argIdx = FieldArgument.[_None] Then
						argIdx = argIdx + 1
						arg.Start = charIndex
						arg.Stop = arg.Start
						
					' ...or extend its location.
					Else
						arg.Stop = arg.Stop + 1
					End If
					
					' Activate nesting...
					If depth = 1 Then
						dfu = dfu + ParsingDefusal.dfuNest
						
						' ...and locate it in the index argument.
						If argIndex = FieldArgument.argIndex Then idxDfu.Start = charIndex
					End If
					
					' Nest deeper into the field.
					depth = depth + 1
					
					' Extend the location of this field.
					expression.Stop = expression.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or parse out of the field...
				Case closeField
					' Unnest out of the field.
					depth = depth - 1
					
					' Extend the location of this field.
					expression.Stop = expression.Stop + 1
					
					' Save (and reset) any argument to the array...
					If argIdx > FieldArgument.[_None] Then
						Expr_Clone arg, args(argIdx)
						Expr_Reset arg
					End If
					
					' ...along with the (field) element.
					status = Fld_Close(e.Field, format := format, expression := expression, args := args, argIdx := argIdx, idxDfu := idxDfu, idxEsc := idxEsc)
					Elm_Clone e, elements(eIdx)
					Elm_Reset e
					
					' Short-circuit for errors.
					If status <> ParsingStatus.stsSuccess Then GoTo EXIT_LOOP
					
					' Advance to the next element.
					eIdx = eIdx + 1
					
					' Reset the global trackers.
					Expr_Reset expression
					Erase args
					argIdx = FieldArgument.[_None]
					Expr_Reset idxDfu
					idxEsc = 0
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or parse to the next argument...
				Case separator
					' Confirm the argument.
					If argIdx = FieldArgument.[_None] Then argIdx = argIdx + 1
					
					' Extend the location of this field.
					expression.Stop = expression.Stop + 1
					
					' Begin a new argument...
					If argIdx < FieldArgument.[_Last] Then
						' Save (and reset) the argument.
						Expr_Clone arg, args(argIdx)
						Expr_Reset arg
						
						' Advance to the next argument.
						argIdx = argIdx + 1
						
						' Locate that argument.
						arg.Stop = charIndex
						arg.Start = arg.Stop + 1
						
					' ...but absorb extra separators into the final argument.
					Else
						' Extend the location of this argument...
						arg.Stop = arg.Stop + 1
						
						' ...along with its contents.
						arg.Syntax = arg.Syntax & char
					End If
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or parse this argument.
				Case Else
					' Confirm the argument...
					If argIdx = FieldArgument.[_None] Then
						argIdx = argIdx + 1
						arg.Start = charIndex
						arg.Stop = arg.Start
						
					' ...or extend its location.
					Else
						arg.Stop = arg.Stop + 1
					End If
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and the contents of this argument.
					arg.Syntax = arg.Syntax & char
					
					' Advance to the next character.
					GoTo NEXT_CHAR
				End Select
			End If
		End Select
		
		
		
	' #############
	' ## Control ##
	' #############
	
	' Advance to the next character.
	NEXT_CHAR:
		
	Next charIndex
	
	
' Escape the loop.
EXIT_LOOP:
	
	' Deactivate error handling.
	On Error GoTo 0
	
	
	
	' ####################
	' ## Interpretation ##
	' ####################
	
	' Short-circuit for any error.
	If status <> ParsingStatus.stsSuccess Then
		Expr_Close expression, format := format
		
	' Handle unresolved syntax: a hanging escape...
	ElseIf Enum_Has(dfu, ParsingDefusal.dfuEscape) Then
		' Return to the final character...
		charIndex = charIndex - 1
		
		' ...and pinpoint the hanging escape.
		expression.Start = charIndex
		expression.Stop = expression.Start
		Expr_Close expression, format := format
		
		' Return the specific status.
		status = ParsingStatus.stsErrorHangingEscape
		
	' ...or an unenclosed quote...
	ElseIf Enum_Has(dfu, ParsingDefusal.dfuQuote) Then
		Select Case e.Kind
		
		' Permit this for regular text...
		Case ElementKind.elmPlain
			' Save this (plaintext) element to the array...
			Elm_Clone e, elements(eIdx)
			
			' ...and report success.
			Expr_Reset expression
			status = ParsingStatus.stsSuccess
			
		' ...but otherwise report the specific error.
		Case Else
			Expr_Close expression, format := format
			status = ParsingStatus.stsErrorUnenclosedQuote
		End Select
		
	' ...or an imbalanced nesting.
	ElseIf Enum_Has(dfu, ParsingDefusal.dfuNest) Or depth > 0 Then
		Expr_Close expression, format := format
		status = ParsingStatus.stsErrorImbalancedNesting
		
	' Otherwise report success in the absence of any issues.
	Else
		' Save any pending (valid) element to the array...
		If e.Kind <> ElementKind.[_Unknown] Then
			Elm_Clone e, elements(eIdx)
		End If
		
		' ...and report success.
		Expr_Reset expression
		status = ParsingStatus.stsSuccess
	End If
	
	
	' Resize the resulting array.
	GoTo RESIZE_ELM
	
	
' Report a generic syntax error.
STX_ERROR:
	' Pinpoint where the error occurred...
	expression.Start = charIndex
	expression.Stop = expression.Start
	Expr_Close expression, format := format
	
	' ...and return the generic status.
	status = ParsingStatus.stsError
	
	
' Resize the array to the elements we actually parsed.
RESIZE_ELM:
	' Locate the last valid element.
	If elements(eIdx).Kind = ElementKind.[_Unknown] Then
		eIdx = eIdx - 1
	End If
	
	' Clear the array if there are no valid elements...
	If eIdx < base Then
		Erase elements
		
	' ...and otherwise truncate it at the last one.
	ElseIf eIdx < eUp Then
		eUp = eIdx
		ReDim Preserve elements(base To eUp)
	End If
End Sub



' #################################
' ## Support | Parsing | Closure ##
' #################################

' Close an expression and record its information.
Private Sub Expr_Close(ByRef expr As ParserExpression, _
	ByRef format As String _
)
	' Clear invalid information...
	If expr.Start <= 0 Then
		Expr_Reset expr
		
	' ...or record a blank...
	ElseIf expr.Start > expr.Stop Then
		expr.Syntax = VBA.vbNullString
		expr.Stop = expr.Start - 1
		
	' ...or record valid syntax.
	Else
		Dim exprLen As Long: exprLen = expr.Stop - expr.Start + 1
		expr.Syntax = VBA.Mid$(format, expr.Start, exprLen)
	End If
End Sub


' Close a field (sub)element and record its information...
Private Function Fld_Close(ByRef fld As ParserField, _
	ByRef format As String, _
	ByRef expression As ParserExpression, _
	ByRef args() As ParserExpression, _
	ByRef argIdx As Long, _
	ByRef idxDfu As ParserExpression, _
	ByRef idxEsc As Long _
) As ParsingStatus
	' Short-circuit for no arguments.
	If argIdx = FieldArgument.[_None] Then
		Fld_Reset fld
		Fld_Close = ParsingStatus.stsSuccess
		Exit Function
	End If
	
	
	' Process each argument...
	Dim arg As ParserExpression
	Fld_Close = ParsingStatus.stsSuccess
	
	' ...except the (trailing) format.
	Dim iTo As Long: iTo = Application.WorksheetFunction.Max(FieldArgument.[_First], argIdx - 1)
	
	Dim i As Long
	For i = FieldArgument.[_First] To iTo
		' Extract the argument by position.
		arg = args(i)
		
		
		Select Case i
		
		' Process the index argument.
		Case FieldArgument.argIndex
			Fld_Close = Fld_CloseIndex(fld, _
				idx := arg, _
				format := format, _
				expression := expression, _
				idxDfu := idxDfu, _
				idxEsc := idxEsc _
			)
		End Select
		
		' Short-circuit for error.
		If Fld_Close <> ParsingStatus.stsSuccess Then GoTo FLD_ERROR
	Next i
	
	
	' Process the (trailing) format.
	If argIdx > FieldArgument.[_First] Then
		' Extract the final argument...
		arg = args(argIdx)
		' Expr_Clone args(argIdx), arg
		
		' ...and process this format.
		Fld_Close = Fld_CloseFormat(fld, _
			fmt := arg, _
			format := format, _
			expression := expression _
		)
	End If
	
	' Short-circuit for error...
	If Fld_Close <> ParsingStatus.stsSuccess Then GoTo FLD_ERROR
	
	' ...and otherwise report success.
	Expr_Reset expression
	Exit Function
	
	
' Report an error when parsing the arguments.
FLD_ERROR:
	Expr_Clone arg, expression
End Function


' ...along with its index argument...
Private Function Fld_CloseIndex(ByRef fld As ParserField, _
	ByRef idx As ParserExpression, _
	ByRef format As String, _
	ByRef expression As ParserExpression, _
	ByRef idxDfu As ParserExpression, _
	ByRef idxEsc As Long _
) As ParsingStatus
	' Define fallback for missing argument.
	Dim noIdx As Variant  ' noIdx = Missing()
	
	' Save the defused syntax...
	Dim dfuSyntax As String: dfuSyntax = idx.Syntax
	
	' ...before recording the original syntax.
	Expr_Close idx, format := format
	
	' Clean the original syntax....
	Dim nLeft As Long, nRight As Long
	Expr_Trim idx, nLeft := nLeft, nRight := nRight
	
	' ...and mirror the defused syntax.
	dfuSyntax = Txt_Crop(dfuSyntax, nLeft := nLeft, nRight := nRight)
	
	' Short-circuit for a missing index.
	If idx.Syntax = VBA.vbNullString Then
		Let fld.Index = noIdx
		Fld_CloseIndex = ParsingStatus.stsSuccess
		Exit Function
	End If
	
	' Check if the index begins with an escape sequence...
	Dim isEsc As Boolean: isEsc = (idxEsc = idx.Start)
	
	' ...or if the index is encapsulated in a single quotation ("...") or nesting ({...}).
	Dim isCap As Boolean: isCap = (idxDfu.Start = idx.Start And idxDfu.Stop = idx.Stop)
	
	' Interpret as an (encapsulated) key...
	If isCap Then
		Let fld.Index = VBA.CStr$(dfuSyntax)
		
	' ...or as an (escaped) key that looks numeric...
	ElseIf isEsc Then
		On Error GoTo IDX_ERROR
		VBA.CLng dfuSyntax
		On Error GoTo 0
		
		' Clean the key while preserving the numeric style ($-1,234.56).
		Let fld.Index = VBA.Trim$(dfuSyntax)
		
	' ...or an integral index.
	Else
		' Defuse a simple (flat) expression with only escapes...
		On Error GoTo IDX_ERROR
		If idxDfu.Start = 0 Then
			Let fld.Index = VBA.CLng(dfuSyntax)
			
		' ...but interpret anything else (deep expressions) as is.
		Else
			Let fld.Index = VBA.CLng(idx.Syntax)
		End If
		On Error GoTo 0
	End If
	
	' Report success.
	Expr_Reset expression
	Fld_CloseIndex = ParsingStatus.stsSuccess
	Exit Function
	
	
' Report the error for an invalid index.
IDX_ERROR:
	Expr_Clone idx, expression
	Fld_CloseIndex = ParsingStatus.stsErrorInvalidIndex
End Function


' ...and its format argument.
Private Function Fld_CloseFormat(ByRef fld As ParserField, _
	ByRef fmt As ParserExpression, _
	ByRef format As String, _
	ByRef expression As ParserExpression _
) As ParsingStatus
	' Define fallback for missing argument.
	Dim noFmt As String
	
	' Record the original syntax...
	Expr_Close fmt, format := format
	
	' ...and short-circuit for a missing format.
	If fmt.Syntax = VBA.vbNullString Then
		Let fmt.Syntax = noFmt
		Fld_CloseFormat = ParsingStatus.stsSuccess
		Exit Function
	End If
	
	' Assign that syntax to the argument...
	fld.Format = fmt.Syntax
	
	' ...and report success.
	Expr_Reset expression
	Fld_CloseFormat = ParsingStatus.stsSuccess
End Function



' ########################
' ## Support | Elements ##
' ########################

' Reset an element.
Private Sub Elm_Reset(ByRef elm As ParserElement)
	Dim reset As ParserElement
	Let elm = reset
End Sub


' Clone one element into another.
Private Sub Elm_Clone(ByRef elm1 As ParserElement, ByRef elm2 As ParserElement)
	Let elm2.Kind   = elm1.Kind
	Let elm2.Plain	= elm1.Plain
	Fld_Clone elm1.Field, elm2.Field
End Sub


' Count the elements returned by Parse().
Private Function Elm_Count(ByRef elms() As ParserExpression) As Long
	On Error GoTo BOUND_ERROR
	Elm_Count = UBound(elms, 1) - LBound(elms, 1) + 1
	Exit Function
	
BOUND_ERROR:
	Elm_Count = 0
End Function



' ######################################
' ## Support | Elements | Expressions ##
' ######################################

' Reset an expression.
Private Sub Expr_Reset(ByRef expr As ParserExpression)
	Dim reset As ParserExpression
	Let expr = reset
End Sub


' Clone one expression into another.
Private Sub Expr_Clone(ByRef expr1 As ParserExpression, ByRef expr2 As ParserExpression)
	Let expr2.Syntax = expr1.Syntax
	Let expr2.Start  = expr1.Start
	Let expr2.Stop   = expr1.Stop
End Sub


' Trim whitespace from an expression.
Private Sub Expr_Trim(ByRef expr As ParserExpression, _
	Optional ByRef nLeft As Long, _
	Optional ByRef nRight As Long _
)
	' Record the initial length...
	Dim n1 As Long, n2 As Long
	n1 = VBA.Len(expr.Syntax)
	
	' ...then trim any trailing whitespace...
	expr.Syntax = VBA.RTrim$(expr.Syntax)
	
	' ...and withdraw the tail.
	n2 = VBA.Len(expr.Syntax)
	nRight = n1 - n2
	expr.Stop = expr.Stop - nRight
	
	
	' Record the remaining length...
	n1 = n2
	
	' ...then trim any leading whitespace...
	expr.Syntax = VBA.LTrim$(expr.Syntax)
	
	' ...and advance the head.
	n2 = VBA.Len(expr.Syntax)
	nLeft = n1 - n2
	expr.Start = expr.Start + nLeft
End Sub



' #################################
' ## Support | Elements | Fields ##
' #################################

' Reset a field.
Private Sub Fld_Reset(ByRef fld As ParserField)
	Dim reset As ParserField
	Let fld = reset
End Sub


' Clone one field (sub)element into another.
Private Sub Fld_Clone(ByRef fld1 As ParserField, ByRef fld2 As ParserField)
	Let fld2.Index    = fld1.Index
	Let fld2.Format   = fld1.Format
End Sub
