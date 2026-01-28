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


' Outcomes of operations like parsing and formatting.
Public Enum sPrinterStatus
	stsSuccess                =    0	' Report success.
	stsError                  = 1000	' Report a general syntax error.
	stsErrorHangingEscape     = 1001	' Report a hanging escape...
	stsErrorUnenclosedQuote   = 1002	' ...or an incomplete quote...
	stsErrorImbalancedNesting = 1003	' ...or an imbalanced nesting...
	stsErrorInvalidIndex      = 1004	' ...or an index that is not an integer.
	stsErrorNonexistentIndex  = 1005	' Report an index that does not exist in the data...
	stsErrorInvalidFormat     = 1006	' ...or a format that is invalid for Format() or Text().
	stsErrorUnknownElement    = 1007	' Report a ParsingElement of unknown kind.
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



' #####################
' ## API | Messaging ##
' #####################

' Embed (formatted) values within a message, sourced from fleXible data.
Public Function xMessage( _
	ByRef format As String, _
	ByRef data As Variant, _
	Optional ByRef lookup As Variant, _
	Optional ByRef default As Variant, _
	Optional ByVal position As PositionKind = PositionKind.posAbsolute, _
	Optional ByVal mode As FormatMode = FormatMode.[_Unknown], _
	Optional ByVal firstDayOfWeek As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbSunday, _
	Optional ByVal firstWeekOfYear As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbFirstJan1, _
	Optional ByVal escape As Variant = SYM_ESC, _
	Optional ByVal openField As String = SYM_FLD_OPEN, _
	Optional ByVal closeField As String = SYM_FLD_CLOSE, _
	Optional ByVal openQuote As String = SYM_QUO_OPEN, _
	Optional ByVal closeQuote As String = SYM_QUO_CLOSE, _
	Optional ByVal separator As String = SYM_SEP _
) As String
	' Short-circuit for blank format.
	If format = VBA.vbNullString Then
		xMessage = VBA.vbNullString
		Exit Function
	End If
	
	
	' Validate the data.
	Dim n As Long, low As Long, up As Long
	Dim isRng As Boolean, ori As Excel.XlRowCol
	CheckData _
		data := data, _
		n := n, _
		low := low, _
		up := up, _
		isRng := isRng, _
		ori := ori
	
	
	' Validate any lookup.
	Dim hasLook As Boolean: hasLook = Not VBA.IsMissing(lookup)
	If hasLook Then
		CheckLookup lookup := lookup
	End If
	
	
	' Parse the message format...
	Dim base As Long: base = 1
	Dim elements() As ParserElement: elements = Parse( _
		format := format, _
		base := base, _
		escape := escape, _
		openField := openField, _
		closeField := closeField, _
		openQuote := openQuote, _
		closeQuote := closeQuote, _
		separator := separator _
	)
	
	' ...and short-circuit for no elements.
	Dim count As Long: count = Elm_Count(elements)
	If count = 0 Then
		xMessage = VBA.vbNullString
		Exit Function
	End If
	
	
	' Assemble the elements into a message.
	Dim eLow As Long: eLow = LBound(elements, 1)
	Dim eUp As Long: eUp = UBound(elements, 1)
	
	Dim hasVal As Boolean, isDfl As Boolean, isAuto As Boolean
	Dim hasDfl As Boolean: hasDfl = Not VBA.IsMissing(default)
	Dim iAuto As Long: iAuto = 0
	Dim iFld As Long: iFld = 0
	Dim e As ParserElement, idx As Variant, pos As PositionKind, val As Variant, fmt As String, out As String
	
	Dim i As Long
	For i = eLow To eUp
		e = elements(i)
		
		Select Case e.Kind
		
		' Simply use plaintext as is...
		Case ElementKind.elmPlain
			out = e.Plain
			
		' ...but format field values for embedding.
		Case ElementKind.elmField
			' Count the field.
			iFld = iFld + 1
			
			' Determine if the field specifies an index.
			isAuto = VBA.IsEmpty(e.Field.Index)
			
			' Default to the next available position...
			If isAuto Then
				iAuto = iAuto + 1
				Let idx = iAuto
				pos = PositionKind.posRelative
				
			' ...unless the field specified an index.
			Else
				Let idx = e.Field.Index
				pos = position
			End If
			
			' Try extracting the value at that index.
			hasVal = GetValue( _
				data := data, _
				idx := idx, _
				n := n, _
				low := low, _
				up := up, _
				isRng := isRng, _
				ori := ori, _
				hasLook := hasLook, _
				lookup := lookup, _
				pos := pos, _
				val := val _
			)
			
			' Handle existing...
			If hasVal Then
				isDfl = False
				
			' ...or nonexisting values.
			Else
				' Use any default...
				If hasDfl Then
					isDfl = True
					Assign val, default
					
				' ...but throw an error otherwise.
				Else
					Err_Index _
						nField := iFld, _
						index := idx, _
						position := pos
				End If
			End If
			
			' Try formatting that value as output.
			fmt = e.Field.Format
			
			On Error GoTo FMT_ERROR
			out = Format2( _
				value := val, _
				format := fmt, _
				mode := mode, _
				firstDayOfWeek := firstDayOfWeek, _
				firstWeekOfYear := firstWeekOfYear _
			)
			On Error GoTo 0
			
		' Throw an error for anything else.
		Case Else
			Err_Element _
				nElement := i, _
				kind := e.Kind
		End Select
		
		' Append the result to the message.
		xMessage = xMessage & out
	Next i
	
	
	' Return the resulting output.
	Exit Function
	
	
' Handle an error when formatting a value. 
FMT_ERROR:
	Err_Format _
		nField := iFld, _
		format := fmt, _
		isDefault := isDfl
End Function


' Embed (formatted) values within a message, sourced from single Values as arguments...
Public Function vMessage( _
	ByRef format As String, _
	ParamArray data() As Variant _
) As String
	vMessage = vxMessage( _
		format := format, _
		data := data _
	)
End Function


' ...and support conversion from those arguments to fleXible data.
Private Function vxMessage( _
	ByRef format As String, _
	ByRef data() As Variant _
) As String
	' Use relative positioning for arguments.
	Dim pos As PositionKind: pos = PositionKind.posRelative
	
	vxMessage = xMessage( _
		format := format, _
		data := data, _
		position := pos _
	)
End Function



' ####################
' ## API | Printing ##
' ####################

' Print such a message to the console: sourced from fleXible data...
Public Function xPrint( _
	ByRef format As String, _
	Optional ByRef data As Variant, _
	Optional ByRef default As Variant, _
	Optional ByVal position As PositionKind = PositionKind.posAbsolute, _
	Optional ByVal mode As FormatMode = FormatMode.[_Unknown], _
	Optional ByVal firstDayOfWeek As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbSunday, _
	Optional ByVal firstWeekOfYear As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbFirstJan1, _
	Optional ByVal escape As Variant = SYM_ESC, _
	Optional ByVal openField As String = SYM_FLD_OPEN, _
	Optional ByVal closeField As String = SYM_FLD_CLOSE, _
	Optional ByVal openQuote As String = SYM_QUO_OPEN, _
	Optional ByVal closeQuote As String = SYM_QUO_CLOSE, _
	Optional ByVal separator As String = SYM_SEP _
) As String
	xPrint = xMessage( _
		format := format, _
		data := data, _
		default := default, _
		position := position, _
		mode := mode, _
		firstDayOfWeek := firstDayOfWeek, _
		firstWeekOfYear := firstWeekOfYear, _
		escape := escape, _
		openField := openField, _
		closeField := closeField, _
		openQuote := openQuote, _
		closeQuote := closeQuote, _
		separator := separator _
	)
	
	Debug.Print xPrint
End Function


' ...or from single Values as arguments.
Public Function vPrint( _
	ByRef format As String, _
	ParamArray data() As Variant _
) As String
	vPrint = vxMessage( _
		format := format, _
		data := data _
	)
	
	Debug.Print vPrint
End Function



' ###################
' ## API | Parsing ##
' ###################

' Parse a format string as an array of syntax elements.
Public Function Parse( _
	ByRef format As String, _
	Optional ByVal base As Long = 1, _
	Optional ByVal escape As Variant = SYM_ESC, _
	Optional ByVal openField As Variant = SYM_FLD_OPEN, _
	Optional ByVal closeField As Variant = SYM_FLD_CLOSE, _
	Optional ByVal openQuote As Variant = SYM_QUO_OPEN, _
	Optional ByVal closeQuote As Variant = SYM_QUO_CLOSE, _
	Optional ByVal separator As Variant = SYM_SEP _
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
	Dim status As sPrinterStatus
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
	If status <> sPrinterStatus.stsSuccess Then
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

' Throw the latest error object.
Private Sub Err_Raise()
	VBA.Err.Raise _
		Number := VBA.Err.Number, _
		Source := VBA.Err.Source, _
		Description := VBA.Err.Description, _
		HelpFile := VBA.Err.HelpFile, _
		HelpContext := VBA.Err.HelpContext
End Sub



' ###############################
' ## Diagnostics | Situational ##
' ###############################

' Throw a parsing error with granular information.
Private Sub Err_Parsing( _
	ByVal status As sPrinterStatus, _
	ByRef expression As ParserExpression, _
	ByVal escape As String, _
	ByVal openField As String, _
	ByVal closeField As String, _
	ByVal openQuote As String, _
	ByVal closeQuote As String, _
	ByVal separator As String _
)
	' Define the horizontal ellipsis: "…"
	#If Mac Then
		Const ETC_SYM As Long = 201
	#Else
		Const ETC_SYM As Long = 133
	#End If
	
	Dim etc As String: etc = Chr2(ETC_SYM)
	
	
	' Describe where the erroneous syntax occurs.
	Dim description As String, position As String, qualifier As String
	Dim startPos As String, stopPos As String
	
	If expression.Start > 0 Then
		startPos = Num_Ordinal(expression.Start)
		stopPos = Num_Ordinal(expression.Stop)
		
		If expression.Start < expression.Stop Then
			qualifier = "somewhere"
			position = "between the " & startPos & " and " & stopPos & " characters"
		ElseIf expression.Start = expression.Stop Then
			position = "at the " & startPos & " character"
		Else
			position = "following the " & stopPos & " character"
		End If
	End If
	
	' Generate a relevant description of the error...
	Select Case status
	Case sPrinterStatus.stsError
		description = "An error occurred when parsing the message format"
		If position <> VBA.vbNullString Then
			If qualifier <> VBA.vbNullString Then position = qualifier & " " & position
			description = description & ", " & position
		End If
		description = description & "."
		
	Case sPrinterStatus.stsErrorHangingEscape
		description = "The message format contains a hanging escape (" & escape & ")"
		If position <> VBA.vbNullString Then description = description & " " & position
		description = description & "."
		
	Case sPrinterStatus.stsErrorUnenclosedQuote
		description = "The message format contains an unenclosed quote (" & openQuote & etc & closeQuote & ")"
		If position <> VBA.vbNullString Then
			If qualifier <> VBA.vbNullString Then position = qualifier & " " & position
			description = description & " " & position
		End If
		description = description & "."
		
	Case sPrinterStatus.stsErrorImbalancedNesting
		description = "The message format contains an imbalanced field nesting (" & openField & etc & closeField & ")"
		If position <> VBA.vbNullString Then
			If qualifier <> VBA.vbNullString Then position = qualifier & " " & position
			description = description & " " & position
		End If
		description = description & "."
		
	Case sPrinterStatus.stsErrorInvalidIndex
		description = "The message format contains an invalid field index"
		If position <> VBA.vbNullString Then description = description & ", " & position
		description = description & ": " & expression.Syntax
		
	' ...but do not raise an unidentified error.
	Case Else
		Exit Sub
	End Select
	
	' Raise the error.
	Err.Raise _
		Number := status, _
		Description := description
End Sub


' Throw an error for an unrecognized .Kind of element.
Private Sub Err_Element( _
	ByVal nElement As Long, _
	ByVal kind As ElementKind _
)
	Dim description As String, kindCode As String
	description = "This " & Num_Ordinal(nElement) & " element from the parser cannot be interpreted"
	description = description & ", because its "".Kind"" (" & Num_Cardinal(kind) & ") is unrecognized."
	
	' Raise the error.
	Err.Raise _
		Number := sPrinterStatus.stsErrorUnknownElement, _
		Description := description
End Sub


' Throw an error for a nonexistent index.
Private Sub Err_Index( _
	ByVal nField As Long, _
	ByRef index As Variant, _
	ByRef position As PositionKind _
)
	' Display the index in detail...
	Dim idxCode As String, idxKind As String, posKind As String
	FormatIndex _
		idx := index, _
		idxCode := idxCode, _
		ord := False, _
		idxKind := idxKind, _
		pos := position, _
		posKind := posKind
	
	' ...with the proper (lower) case.
	idxCode = VBA.LCase$(idxCode)
	idxKind = VBA.LCase$(idxKind)
	posKind = VBA.LCase$(posKind)
	
	' Generate a relevant description of the error.
	Dim description As String
	description = "This"
	If posKind <> VBA.vbNullString Then description = description & " (" & posKind & ")"
	description = description & " " & idxKind & " does not exist in the data"
	description = description & ", as given in the " & Num_Ordinal(nField) & " field of the message format"
	description = description & ": " & idxCode
	
	' Raise the error.
	Err.Raise _
		Number := sPrinterStatus.stsErrorNonexistentIndex, _
		Description := description
End Sub


' Throw an error for an invalid format.
Private Sub Err_Format( _
	ByVal nField As Long, _
	ByRef format As String, _
	Optional ByVal isDefault As Boolean = False _
)
	' Generate a relevant description of the error.
	Dim description As String, valKind As String
	If isDefault Then valKind = "default"
	
	description = "The"
	If valKind <> VBA.vbNullString Then description = description & " (" & valKind & ")"
	description = description & " value cannot be displayed in this format code"
	description = description & ", as given in the " & Num_Ordinal(nField) & " field of the message format"
	description = description & ": " & format
	
	' Raise the error.
	Err.Raise _
		Number := sPrinterStatus.stsErrorInvalidFormat, _
		Description := description
End Sub


' Throw an error for a blank parsing symbol.
Private Sub Err_BlankSym()
	' Detail the error: invalid procedure call or argument.
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
	' Define the error: invalid procedure call or argument.
	Const ERR_NUM As Long = 5
	
	' Define the horizontal ellipsis: "…"
	#If Mac Then
		Const ETC_SYM As Long = 201
	#Else
		Const ETC_SYM As Long = 133
	#End If
	
	Dim etc As String: etc = Chr2(ETC_SYM)
	
	
	' Generate a relevant description of the error.
	Dim description As String
	description = "The same formatting symbol (""" & sym & """) may not be used twice"
	description = description & "; " & "aside from quotes (" & openQuote & etc & closeQuote & ") if you so specify"
	description = description & "."
	
	' Raise the error.
	Err.Raise _
		Number := ERR_NUM, _
		Description := description
End Sub


' Throw an error for data of the wrong structure.
Private Sub Err_Data()
	' Define the error: type mismatch.
	Const ERR_NUM As Long = 13
	
	' Specify formatting.
	Const DESC_IND As String = "  "
	
	
	' Generate a relevant description of the error.
	Dim description As String: description = "The source data must be one of the following"
	Dim types As Variant: types = Array( _
		"A unidimensional (1D) array.", _
		"A Range of cells, in a single row or single column.", _
		"An (initialized) object with a "".Count"" property and default member." _
	)
	description = description & ":" & VBA.vbNewLine & Txt_List(types, indent := DESC_IND, separator := VBA.vbNewLine)
	
	' Raise the error.
	Err.Raise _
		Number := ERR_NUM, _
		Description := description
End Sub


' Throw an error for a lookup of the wrong structure.
Private Sub Err_Lookup()
	' Define the error: type mismatch.
	Const ERR_NUM As Long = 13
	
	' Specify formatting.
	Const DESC_IND As String = "  "
	
	
	' Generate a relevant description of the error.
	Dim description As String: description = "The lookup data must be one of the following"
	Dim types As Variant: types = Array( _
		"A unidimensional (1D) array.", _
		"A Range of cells, in a single row or single column." _
	)
	description = description & ":" & VBA.vbNewLine & Txt_List(types, indent := DESC_IND, separator := VBA.vbNewLine)
	
	' Raise the error.
	Err.Raise _
		Number := ERR_NUM, _
		Description := description
End Sub



' ###############
' ## Utilities ##
' ###############

' Assign any value (scalar or objective) to a variable.
Public Sub Assign(ByRef var As Variant, ByRef val As Variant)
	If VBA.IsObject(val) Then
		Set var = val
	Else
		Let var = val
	End If
End Sub


' Test if a combination (dfuNest + dfuEscape) includes a particular enumeration (dfuEscape).
Public Function Enum_Has(ByRef enum1 As Long, ByRef enum2 As Long) As Boolean
	Enum_Has = enum1 And enum2
End Function



' #########################
' ## Utilities | Numbers ##
' #########################

' Display a cardinal integer: 1,234
Public Function Num_Cardinal(ByVal num As Long) As String
	Const CARD_FMT As String = "#,##0"
	
	Num_Cardinal = VBA.Format(num, Format := CARD_FMT)
End Function


' Display an ordinal integer: 1,234th
Public Function Num_Ordinal(ByVal num As Long) As String
	Const NUM_BASE As Integer = 10
	
	' Determine the proper suffix...
	Dim sfx As String
	Select Case (num Mod NUM_BASE)
	Case 1:		sfx = "st"
	Case 2:		sfx = "nd"
	Case 3:		sfx = "rd"
	Case Else:	sfx = "th"
	End Select
	
	' ...and append it to the cardinal.
	Num_Ordinal = Num_Cardinal(num) & sfx
End Function



' ########################
' ## Utilities | Arrays ##
' ########################

' Get the length (along a dimension) of an array.
Public Function Arr_Length(ByRef arr As Variant, _
	Optional ByVal dimension As Long = 1 _
) As Long
	' Subscript out of range.
	Const EMPTY_ERR As Long = 9
	
	On Error GoTo BOUND_ERROR
	Arr_Length = UBound(arr, dimension) - LBound(arr, dimension) + 1
	Exit Function
	
	
' Handle an empty array.
BOUND_ERROR:
	Select Case VBA.Err.Number
	Case EMPTY_ERR:	Arr_Length = 0
	Case Else:	Err_Raise
	End Select
End Function


' Get the "rank" of an array: the count of its dimensions.
Public Function Arr_Rank(ByRef arr As Variant) As Long
	' Subscript out of range.
	Const EMPTY_ERR As Long = 9
	
	Dim tst As Long
	Arr_Rank = 0
	
	On Error GoTo BOUND_ERROR
	Do While True
		Arr_Rank = Arr_Rank + 1
		tst = UBound(arr, Arr_Rank)
	Loop
	
	
' Handle the rank.
BOUND_ERROR:
	Select Case VBA.Err.Number
	Case EMPTY_ERR:	Arr_Rank = Arr_Rank - 1
	Case Else:	Err_Raise
	End Select
End Function



' ######################
' ## Utilities | Text ##
' ######################

' Get a Unicode character, without sporadic corruption on Mac.
Public Function Chr2(ByVal charcode As Long) As String
	#If Mac Then
		Chr2 = VBA.Chr(charcode)
	#Else
		Chr2 = VBA.ChrW(charcode)
	#End If
End Function


' Remove characters from the end(s) of a string.
Public Function Txt_Crop(ByVal txt As String, _
	Optional ByVal nLeft As Long = 0, _
	Optional ByVal nRight As Long = 0 _
) As String
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


' Display phrases of text within a list.
Public Function Txt_List(ByRef txts As Variant, _
	Optional ByVal symbol As String = "-", _
	Optional ByVal indent As String = VBA.vbTab, _
	Optional ByVal offset As String = " ", _
	Optional ByVal separator As String = VBA.vbNewLine _
) As String
	Dim pfx As String: pfx = indent & symbol & offset
	Dim sep As String: sep = separator & pfx
	
	If Arr_Length(txts) = 0 Then
		Exit Function
	Else
		Txt_List = VBA.Join(sourcearray := txts, delimiter := sep)
		Txt_List = pfx & Txt_List
	End If
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

' Display the index for a field...
Private Sub FormatIndex( _
	ByRef idx As Variant, _
	Optional ByRef idxCode As String, _
	Optional ByVal ord As Boolean = False, _
	Optional ByRef idxKind As String, _
	Optional ByVal pos As PositionKind, _
	Optional ByRef posKind As String _
)
	' Define how a key is displayed.
	Const KEY_OPEN As String = """"
	Const KEY_CLOSE As String = """"
	
	
	Select Case VBA.VarType(idx)
	Case VBA.VbVarType.vbLong
		If ord Then
			idxCode = Num_Ordinal(idx)
		Else
			idxCode = Num_Cardinal(idx)
		End If
		
		idxKind = "Position"
		
		Select Case pos
		Case PositionKind.posAbsolute: posKind = "Absolute"
		Case PositionKind.posRelative: posKind = "Relative"
		End Select
		
	Case VBA.VbVarType.vbString
		idxCode = KEY_OPEN & idx & KEY_CLOSE
		idxKind = "Key"
	End Select
End Sub



' ##########################
' ## Support | Validation ##
' ##########################

' ####################################
' ## Support | Validation | Symbols ##
' ####################################

' Validate input (text or code) for a parsing symbol.
Private Sub CheckSym(ByRef x As Variant)
	' Extract the first character from a string...
	If VBA.VarType(x) = VBA.VbVarType.vbString Then
		x = VBA.Left$(x, 1)
		
	' ...or convert a code into its character.
	Else
		x = Chr2(x)
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



' #################################
' ## Support | Validation | Data ##
' #################################

' Validate input for the data structure whose values are embedded.
Private Sub CheckData( _
	ByRef data As Variant, _
	Optional ByRef n As Long, _
	Optional ByRef low As Long, _
	Optional ByRef up As Long, _
	Optional ByRef isRng As Boolean, _
	Optional ByRef ori As Excel.XlRowCol _
)
	' Define an unspecified orientation.
	Const NO_ORI As Long = 0
	
	
	' By default the data is not a Range with any orientation.
	isRng = False
	ori = NO_ORI
	
	' Examine an object...
	If VBA.IsObject(data) Then
		' Short-circuit for an uninitialized object...
		If data Is Nothing Then GoTo DATA_ERROR
		
		' ...but otherwise check an initialized object.
		CheckObject _
			obj := data, _
			n := n, _
			low := low, _
			up := up, _
			isRng := isRng, _
			ori := ori
		
	' ...or an array...
	ElseIf VBA.IsArray(data) Then
		On Error GoTo DATA_ERROR
		CheckArray _
			arr := data, _
			n := n, _
			low := low, _
			up := up
		On Error GoTo 0
		
	' ...but throw an error for anything else.
	Else
		GoTo DATA_ERROR
	End If
	
	' Conclude validation successfully.
	Exit Sub
	
	
' Report an error for an invalid structure.
DATA_ERROR:
	Err_Data
End Sub


' Validate input for names, with which to "look up" the original data.
Private Sub CheckLookup(ByRef lookup As Variant)
	' Examine an object...
	If VBA.IsObject(lookup) Then
		' Short-circuit for an uninitialized object...
		If lookup Is Nothing Then GoTo LOOK_ERROR
		
		' ...or for anything other than a Range.
		If Not TypeOf lookup Is Range Then GoTo LOOK_ERROR
		
		' Check a Range specifically.
		On Error GoTo LOOK_ERROR
		CheckRange rng := lookup
		On Error GoTo 0
		
	' ...or an array...
	ElseIf VBA.IsArray(lookup) Then
		On Error GoTo LOOK_ERROR
		CheckArray arr := lookup
		On Error GoTo 0
		
	' ...but throw an error for anything else.
	Else
		GoTo LOOK_ERROR
	End If
	
	' Conclude validation successfully.
	Exit Sub
	
	
' Report an error for invalid structure.
LOOK_ERROR:
	Err_Lookup
End Sub


' Validate an object as input.
Private Sub CheckObject( _
	ByVal obj As Object, _
	Optional ByRef n As Long, _
	Optional ByRef low As Long, _
	Optional ByRef up As Long, _
	Optional ByRef isRng As Boolean, _
	Optional ByRef ori As Excel.XlRowCol _
)
	' Define the error: type mismatch.
	Const ERR_NUM As Long = 13
	
	
	' The base used by objects like Collections.
	Const OBJ_BASE As Long = 1
	
	
	' Check a Range specifically...
	isRng = TypeOf obj Is Range
	If isRng Then
		On Error GoTo OBJ_ERROR
		CheckRange _
			rng := obj, _
			n := n, _
			low := low, _
			up := up, _
			ori := ori
		On Error GoTo 0
			
	' ...or some other object.
	Else
		On Error GoTo OBJ_ERROR
		n = obj.Count
		On Error GoTo 0
		
		If n > 0 Then
			low = OBJ_BASE
			up = low + n - 1
		Else
			low = OBJ_BASE - 1
			up = low
		End If
	End If
	
	' Conclude validation successfully.
	Exit Sub
	
	
' Report an error for invalid structure.
OBJ_ERROR:
	Err.Raise Number := ERR_NUM
End Sub


' Validate a (1D) Range as input.
Private Sub CheckRange( _
	ByVal rng As Range, _
	Optional ByRef n As Long, _
	Optional ByRef low As Long, _
	Optional ByRef up As Long, _
	Optional ByRef ori As Excel.XlRowCol _
)
	' Define the error: type mismatch.
	Const ERR_NUM As Long = 13
	
	
	' The base used by Ranges of cells.
	Const RNG_BASE As Long = 1
	
	' An unspecified orientation.
	Const NO_ORI As Long = 0
	
	
	' Measure the dimensions of the Range.
	Dim nRows As Double: nRows = rng.Rows.CountLarge
	Dim nCols As Double: nCols = rng.Columns.CountLarge
	
	' Throw an error for a rectangular (2D) area.
	If nRows > 1 And nCols > 1 Then
		GoTo RNG_ERROR
		
	' Handle a single row...
	ElseIf nCols > 1 Then
		n = nCols
		ori = Excel.XlRowCol.xlRows
		
	' ...or a single column...
	ElseIf nRows > 1 Then
		n = nRows
		ori = Excel.XlRowCol.xlColumns
		
	' ...or a single cell.
	Else
		n = 1
		ori = NO_ORI
	End If
	
	' Record remaining information.
	low = RNG_BASE
	up = low + n - 1
	
	' Conclude validation successfully.
	Exit Sub
	
	
' Report an error for invalid structure.
RNG_ERROR:
	Err.Raise Number := ERR_NUM
End Sub


' Validate a (1D) array as input.
Private Sub CheckArray( _
	ByRef arr As Variant, _
	Optional ByRef n As Long, _
	Optional ByRef low As Long, _
	Optional ByRef up As Long _
)
	' Define the error: type mismatch.
	Const ERR_NUM As Long = 13
	
	
	' Ensure the array is a (1D) vector...
	Dim rnk As Long: rnk = Arr_Rank(arr)
	If rnk <> 1 Then GoTo ARR_ERROR
	
	' ...that is not empty.
	n = Arr_Length(arr, dimension := 1)
	If n <= 0 Then GoTo ARR_ERROR
	
	' Record the bounds...
	low = LBound(arr, 1)
	up = UBound(arr, 1)
	
	' ...and conclude validation successfully.
	Exit Sub
	
	
' Report an error for invalid structure.
ARR_ERROR:
	Err.Raise Number := ERR_NUM
End Sub



' ##########################
' ## Support | Extraction ##
' ##########################

' Extract a value from a data structure.
Private Function GetValue( _
	ByRef data As Variant, _
	ByVal idx As Variant, _
	ByVal n As Long, _
	ByVal low As Long, _
	ByVal up As Long, _
	ByVal isRng As Boolean, _
	ByVal ori As Excel.XlRowCol, _
	Optional ByRef hasLook As Boolean = False, _
	Optional ByRef lookup As Variant, _
	Optional ByVal pos As PositionKind = PositionKind.posAbsolute, _
	Optional ByRef val As Variant _
) As Boolean
	' Short-circuit for no data.
	If n = 0 Then GoTo VAL_ERROR
	
	' Identify the type of index...
	Dim idxType As VBA.VbVarType: idxType = VBA.VarType(idx)
	Dim isPos As Boolean: isPos = (idxType = VBA.VbVarType.vbLong)
	Dim isKey As Boolean: isKey = (idxType = VBA.VbVarType.vbString)
	
	' Handle textual keys...
	If isKey Then
		' Look up the key where appropriate...
		If hasLook Then
			Dim hasLoc As Boolean, loc As Long
			hasLoc = LookupKey( _
				lookup := lookup, _
				key := idx, _
				pos := loc _
			)
			
			' ...and short-circuit for no match.
			If Not hasLoc Then GoTo VAL_ERROR
			
			' Use relative positioning for lookups against the source data.
			idx = loc
			isPos = True
			pos = PositionKind.posRelative
		End If
	End If
	
	' ...and numeric positions.
	If isPos Then
		' Interpret those which are relative.
		If pos = PositionKind.posRelative Then
			' Report failure for those out of bounds...
			If idx = 0 Or idx > n Then
				GoTo VAL_ERROR
				
			' ...but otherwise count from the beginning...
			ElseIf idx > 0 Then
				idx = low + idx - 1
				
			' ...or from the end.
			ElseIf idx < 0 Then
				idx = up + idx + 1
			End If
		End If
		
		' Short-circuit for a position that is out of bounds.
		If idx < low Or idx > up Then GoTo VAL_ERROR
	End If
	
	' Extract the value from a Range...
	If isRng Then
		' Short-circuit without a positional index.
		If Not isPos Then GoTo VAL_ERROR
		
		GetValue = GetRangeValue( _
			rng := data, _
			pos := idx, _
			ori := ori, _
			val := val _
		)
		Exit Function
		
	' ...or from something else.
	Else
		On Error GoTo VAL_ERROR
		Assign val, data(idx)
		On Error GoTo 0
	End If
	
	' Report success.
	GetValue = True
	Exit Function
	
	
' Report a nonexistent value.
VAL_ERROR:
	GetValue = False
End Function


' Extract a value from a Range.
Private Function GetRangeValue( _
	ByVal rng As Range, _
	ByVal pos As Long, _
	ByVal ori As Excel.XlRowCol, _
	Optional ByRef val As Variant _
) As Boolean
	' The base used by Ranges of cells.
	Const RNG_BASE As Long = 1
	
	' An unspecified orientation.
	Const NO_ORI As Long = 0
	
	
	On Error GoTo VAL_ERROR
	Select Case ori
	
	' Extract from the first row...
	Case Excel.XlRowCol.xlRows
		Assign val, rng.Cells(RNG_BASE, pos).Value
		
	' ...or from the first column...
	Case Excel.XlRowCol.xlColumns
		Assign val, rng.Cells(pos, RNG_BASE).Value
		
	' ...or from the first (single) cell.
	Case NO_ORI
		' Short-circuit for any position other than the first.
		If pos <> RNG_BASE Then GoTo VAL_ERROR
		
		Assign val, rng.Value
		
	' Short-circuit for any other orientation.
	Case Else
		GoTo VAL_ERROR
	End Select
	On Error GoTo 0
	
	
	' Report success.
	GetRangeValue = True
	Exit Function
	
	
' Report a nonexistent value.
VAL_ERROR:
	GetRangeValue = False
End Function


' Locate a key within lookup data.
Private Function LookupKey( _
	ByRef lookup As Variant, _
	ByVal key As String, _
	Optional ByRef pos As Long _
) As Boolean
	' Search forwards for an exact match.
	Const MATCH_MODE As Long = 0
	Const SEARCH_MODE As Long = 1
	
	
	' Perform the search to record the location...
	On Error GoTo LOOK_ERROR
	pos = Application.WorksheetFunction.XMatch( _
		Arg1 := key, _
		Arg2 := lookup, _
		Arg3 := MATCH_MODE, _
		Arg4 := SEARCH_MODE _
	)
	On Error GoTo 0
	
	' ...and report success.
	LookupKey = True
	Exit Function
	
	
' Report a nonexistent key.
LOOK_ERROR:
	LookupKey = False
End Function



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
	Optional ByRef status As sPrinterStatus _
)
	' ###########
	' ## Setup ##
	' ###########
	
	' Default to success.
	status = sPrinterStatus.stsSuccess
	
	' Record the format length.
	Dim fmtLen As Long: fmtLen = VBA.Len(format)
	
	' Short-circuit for unformatted input.
	If fmtLen = 0 Then
		Erase elements
		Expr_Reset expression
		status = sPrinterStatus.stsSuccess
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
					
					' Locate escaping in any index argument.
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
					If status <> sPrinterStatus.stsSuccess Then GoTo EXIT_LOOP
					
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
	If status <> sPrinterStatus.stsSuccess Then
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
		status = sPrinterStatus.stsErrorHangingEscape
		
	' ...or an unenclosed quote...
	ElseIf Enum_Has(dfu, ParsingDefusal.dfuQuote) Then
		Select Case e.Kind
		
		' Permit this for regular text...
		Case ElementKind.elmPlain
			' Save this (plaintext) element to the array...
			Elm_Clone e, elements(eIdx)
			
			' ...and report success.
			Expr_Reset expression
			status = sPrinterStatus.stsSuccess
			
		' ...but otherwise report the specific error.
		Case Else
			Expr_Close expression, format := format
			status = sPrinterStatus.stsErrorUnenclosedQuote
		End Select
		
	' ...or an imbalanced nesting.
	ElseIf Enum_Has(dfu, ParsingDefusal.dfuNest) Or depth > 0 Then
		Expr_Close expression, format := format
		status = sPrinterStatus.stsErrorImbalancedNesting
		
	' Otherwise report success in the absence of any issues.
	Else
		' Save any pending (valid) element to the array...
		If e.Kind <> ElementKind.[_Unknown] Then
			Elm_Clone e, elements(eIdx)
		End If
		
		' ...and report success.
		Expr_Reset expression
		status = sPrinterStatus.stsSuccess
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
	status = sPrinterStatus.stsError
	
	
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
) As sPrinterStatus
	' Short-circuit for no arguments.
	If argIdx = FieldArgument.[_None] Then
		Fld_Reset fld
		Fld_Close = sPrinterStatus.stsSuccess
		Exit Function
	End If
	
	
	' Process each argument...
	Dim arg As ParserExpression
	Fld_Close = sPrinterStatus.stsSuccess
	
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
		If Fld_Close <> sPrinterStatus.stsSuccess Then GoTo FLD_ERROR
	Next i
	
	
	' Process the (trailing) format.
	If argIdx > FieldArgument.[_First] Then
		' Extract the final argument...
		arg = args(argIdx)
		
		' ...and process this format.
		Fld_Close = Fld_CloseFormat(fld, _
			fmt := arg, _
			format := format, _
			expression := expression _
		)
	End If
	
	' Short-circuit for error...
	If Fld_Close <> sPrinterStatus.stsSuccess Then GoTo FLD_ERROR
	
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
) As sPrinterStatus
	' Define fallback for missing argument.
	Dim noIdx As Variant
	
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
		Fld_CloseIndex = sPrinterStatus.stsSuccess
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
	Fld_CloseIndex = sPrinterStatus.stsSuccess
	Exit Function
	
	
' Report the error for an invalid index.
IDX_ERROR:
	Expr_Clone idx, expression
	Fld_CloseIndex = sPrinterStatus.stsErrorInvalidIndex
End Function


' ...and its format argument.
Private Function Fld_CloseFormat(ByRef fld As ParserField, _
	ByRef fmt As ParserExpression, _
	ByRef format As String, _
	ByRef expression As ParserExpression _
) As sPrinterStatus
	' Define fallback for missing argument.
	Dim noFmt As String
	
	' Record the original syntax...
	Expr_Close fmt, format := format
	
	' ...and short-circuit for a missing format.
	If fmt.Syntax = VBA.vbNullString Then
		Let fmt.Syntax = noFmt
		Fld_CloseFormat = sPrinterStatus.stsSuccess
		Exit Function
	End If
	
	' Assign that syntax to the argument...
	fld.Format = fmt.Syntax
	
	' ...and report success.
	Expr_Reset expression
	Fld_CloseFormat = sPrinterStatus.stsSuccess
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
Private Function Elm_Count(ByRef elms() As ParserElement) As Long
	' Subscript out of range.
	Const EMPTY_ERR As Long = 9
	
	On Error GoTo BOUND_ERROR
	Elm_Count = UBound(elms, 1) - LBound(elms, 1) + 1
	Exit Function
	
BOUND_ERROR:
	Select Case VBA.Err.Number
		Case EMPTY_ERR:	Elm_Count = 0
		Case Else:	Err_Raise
	End Select
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
