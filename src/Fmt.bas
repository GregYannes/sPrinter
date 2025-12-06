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
' 	[_Unknown] = 0	' Uninitialized.
	fmtVbFormat	' The Format() function in VBA.
	fmtXlText	' The Text() function in Excel.
End Enum


' ' Syntax for parsing.
' Private Enum ParsingSymbol
' '	=============	====	  =============		=========
' '	Label		Code	  Name			Character
' '	=============	====	  =============		=========
' 	symEscape     =	  92	' Backslash		\
' 	symOpenField  =	 173	' Opening brace		{
' 	symCloseField =	 175	' Closing brace		}
' 	symopenQuote  =	  34	' Double quotes		"
' 	symCloseQuote =	  34	' Double quotes		"
' 	symSeparator  =	  58	' Colon			:
' End Enum


' Outcomes of parsing.
Public Enum ParsingStatus
	stsSuccess              =    0	' Report success.
	stsError                = 1000	' Report a general syntax error.
	stsErrorHangingEscape   = 1001	' Report a hanging escape...
	stsErrorUnenclosedField = 1002	' ...or an incomplete field...
	stsErrorUnenclosedQuote = 1003	' ...or an incomplete quote...
	stsErrorInvalidIndex    = 1004	' ...or an index that is not an integer.
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
' 	argPosition				' How to interpret a negative index.
' 	argDay1					' The first day of the week, passed to Format() as "FirstDayOfWeek".
' 	argWeek1				' The first week of the year, passed to Format() as "FirstWeekOfYear".
	argFormat				' The formatting applied to the value.
	[_All]					' All arguments.
	
	[_First] = FieldArgument.[_None] + 1	' The first argument.
	[_Last]  = FieldArgument.[_All] - 1	' The last argument.
End Enum


' ' Ways to interpret (negative) positional indices.
' Public Enum PositionKind
' 	[_Unknown] = 0	' Uninitialized.
' 	posAbsolute	' Negative index (-1) is extracted directly...
' 	posRelative	' ...or measured (1st) from the end.
' End Enum



' ###########
' ## Types ##
' ###########

' An expression for parsing.
Public Type ParserExpression
	Syntax As String	' The syntax that was parsed to define this expression.
	Start As Long		' Where that syntax begins in the original code...
	Stop AS Long		' ...and where it ends.
End Type


' Element for parsing a field embedded in formatting.
Public Type ParserField
	Index As Variant		' Any index to extract the value...
' 	Position As PositionKind	' ...and how to interpret that index.
' 	Day1 As VBA.VbDayOfWeek		' The convention for weekdays...
' 	Week1 As VBA.VbFirstWeekOfYear	' ...and calendar weeks...
	Format As String		' ...for any format applied to the value.
End Type


' Elements into which formats are parsed.
Public Type ParserElement
	Kind As ElementKind		' The subtype which extends this element:
	Plain as String			'   - Plain text which displays literally...
	Field As ParserField		'   - ...or a field which embeds a value.
End Type



' #########
' ## API ##
' #########

' .
Public Function Parse( _
	ByRef format As String, _
	ByRef elements() As ParserElement, _
	Optional ByRef expression As ParserExpression, _
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
		Erase elements
		Expr_Reset expression
		Parse = ParsingStatus.stsSuccess
		Exit Function
	End If
	
	
	' Size to accommodate all (possible) elements.
	Dim eLen As Long: eLen = VBA.Int(fmtLen / 2) + 1
	Dim eUp As Long: eUp = base + eLen - 1
	ReDim elements(base To eUp)
	
	
	' Track the current context for parsing...
	Dim dfu As ParsingDefusal: dfu = ParsingDefusal.[_Off]
	Dim depth As Long: depth = 0
	Dim endStatus As ParsingStatus: endStatus = ParsingStatus.stsSuccess
	
	' ...and the current element...
	Dim eIdx As Long: eIdx = base
	Dim e As ParserElement
	Expr_Reset expression
	
	' ...and the current (field) argument...
	Dim args(FieldArgument.[_First] To FieldArgument.[_Last]) As ParserExpression
	Dim argIdx As FieldArgument: argIdx = FieldArgument.[_None]
	Dim arg As ParserExpression
	Dim nDfu As Long: nDfu = 0
	Dim idxEsc As Boolean: idxEsc = False
	
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
		
	' Revisit the character.
	SAME_CHAR:
		' Interpret this character in context.
		Select Case e.Kind
		
		
		
		' ##############
		' ## Inactive ##
		' ##############
		
		Case ElementKind.[_Unknown]
			' ...
			
			
			Select Case char
			
			' Parse into a field...
			Case openField
				' ...
				
			' ...or interpret as text.
			Case Else
				' ...
			End Select
		End If
		
		
		
		' ################
		' ## Plain Text ##
		' ################
		
		Case ElementKind.elmPlain
			' Escape a literal character...
			If Enum_Has(ParsingDefusal.dfuEscape) Then
				' ...
				
			' ...or quote "inert" text...
			ElseIf Enum_Has(dfu, ParsingDefusal.dfuQuote) Then
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					' ...
					
				' ...or continue quoting.
				Case Else
					' ...
				End Select
				
			' ...or parse "active" expressions.
			Case Else
				Select Case char
				
				' Escape the next character...
				Case escape
					' ...
					
				' ...or quote the next characters...
				Case openQuote
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
		
		Case ElementKind.elmField
			' Escape a literal character...
			If Enum_Has(dfu, ParsingDefusal.dfuEscape) Then
				' ...
				
			' ...or quote "inert" text...
			ElseIf Enum_Has(dfu, ParsingDefusal.dfuQuote) Then
				Select Case char
				
				' Terminate the quote...
				Case closeQuote
					' ...
					
				' ...or continue quoting.
				Case Else
					' ...
				End Select
				
			' ...or nest expressions...
			ElseIf Enum_Has(dfu, ParsingDefusal.dfuNest) Then
				Select case char
				
				' Escape the next character...
				Case escape
					' ...
					
				' ...or quote the next characters...
				Case openQuote
					' ...
					
				' ...or nest deeper...
				Case openField
					' ...
					
				' ...or unnest shallower...
				Case closeField
					' ...
					
				' ...or display literally.
				Case Else
					' ...
				End Select
				
			' ...or parse "active" expressions.
			Case Else
				Select Case char
				
				' Escape the next character...
				Case escape
					' ...
					
				' ...or quote the next characters...
				Case openQuote
					' ...
					
				' ...or nest the next syntax...
				Case openField
					' ...
					
				' ...or parse out of the field...
				Case closeField
					' ...
					
				' ...or parse to the next argument...
				Case separator
					' ...
					
				' ...or parse this argument.
				Case Else
					' ...
				End Select
			End Select
		End Select
		
		
		
	' #############
	' ## Control ##
	' #############
	
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
			nDfu := nDfu, _
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
		If depth <> 0 Then
			Parse = ParsingStatus.stsErrorUnenclosedField
			
		' ...or a index of the wrong type...
		ElseIf endStatus = ParsingStatus.stsErrorInvalidIndex Then
			Parse = ParsingStatus.stsErrorInvalidIndex
			
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



' ###############
' ## Utilities ##
' ###############

' Test if a combination (dfuNest + dfuEscape) includes a particular enumeration (dfuEscape).
Public Function Enum_Has(ByRef enum1 As Long, ByRef enum2 As Long) As Boolean
	Enum_Is = enum1 And enum2
End Sub



' #############
' ## Support ##
' #############

' #######################
' ## Support | Parsing ##
' #######################

' ' Reset any global trackers.
' Private Sub Reset( _
' 	Optional ByRef dfu As ParsingDefusal, _
' 	Optional ByRef depth As Long, _
' 	Optional ByRef eIdx As Long, _
' 	Optional ByRef e As ParserElement, _
' 	Optional ByRef char As String, _
' 	Optional ByRef nDfu As Long, _
' 	Optional ByRef idxEsc As Boolean, _
' 	Optional ByRef endStatus As ParsingStatus _
' )
' 	dfu = ParsingDefusal.[_Off]
' 	depth = 0
' 	eIdx = 0
' 	Elm_Reset e
' 	char = VBA.vbNullString
' 	nDfu = 0
' 	idxEsc = False
' 	endStatus = ParsingStatus.stsSuccess
' End Sub


' ' Save an element.
' Private Function Save( _
' 	ByRef format As String, _
' 	ByRef elements As ParserElement(), _
' 	ByRef eIdx As Long, _
' 	ByRef e As ParserElement, _
' 	ByRef nDfu As Long, _
' 	ByRef idxEsc As Boolean _
' ) As ParsingStatus
' 	Save = Elm_Close(e, format := format, nDfu := nDfu, idxEsc := idxEsc)
' 	Elm_Clone e, elements(eIdx)
' End Function


' ' Interpret specifiers for field arguments.
' Private Function Fld_Spec( _
' 	ByVal spec As String, _
' 	ByVal arg As FieldArgument, _
' 	ByRef exists As Boolean _
' ) As Long
' 	' spec = VBA.Trim(spec)
' 	spec = VBA.UCase(spec)
' 	
' 	' 		============	==========	======================
' 	' 		Abbreviation	Name      	Enumeration
' 	' 		============	==========	======================
' 	Select Case arg
' 	
' 	' ' How to interpret a negative index.
' 	' Case FieldArgument.argPosition
' 	' 	Select Case spec
' 	' 	' 	============	==========	======================
' 	' ' 	Case	"?", 		"UNKNOWN",	"[_UNKNOWN]":		Fld_Spec = PositionKind.[_Unknown]
' 	' 	Case	"ABS",		"ABSOLUTE",	"POSABSOLUTE":		Fld_Spec = PositionKind.posAbsolute
' 	' 	Case	"REL",		"RELATIVE",	"POSRELATIVE":		Fld_Spec = PositionKind.posRelative
' 		' 	============	==========	======================
' 		Case Else:												GoTo NO_MATCH
' 	' 	Case Else:							Fld_Spec = PositionKind.[_Unknown]:	GoTo NO_MATCH
' 	' 	End Select
' 		
' 	' ' The "FirstDayOfWeek" for Format().
' 	' Case FieldArgument.argDay1
' 	' 	Select Case spec
' 	' 	' 	============	==========	======================
' 	' 	Case	"SYS",		"SYSTEM",	"VBUSESYSTEMDAYOFWEEK":	Fld_Spec = VBA.VbDayOfWeek.vbUseSystemDayOfWeek
' 	' 	Case	"SUN",		"SUNDAY",	"VBSUNDAY":		Fld_Spec = VBA.VbDayOfWeek.vbSunday
' 	' 	Case	"MON",		"MONDAY",	"VBMONDAY":		Fld_Spec = VBA.VbDayOfWeek.vbMonday
' 	' 	Case	"TUE",		"TUESDAY",	"VBTUESDAY":		Fld_Spec = VBA.VbDayOfWeek.vbTuesday
' 	' 	Case	"WED",		"WEDNESDAY",	"VBWEDNESDAY":		Fld_Spec = VBA.VbDayOfWeek.vbWednesday
' 	' 	Case	"THU",		"THURSDAY",	"VBTHURSDAY":		Fld_Spec = VBA.VbDayOfWeek.vbThursday
' 	' 	Case	"FRI",		"FRIDAY",	"VBFRIDAY":		Fld_Spec = VBA.VbDayOfWeek.vbFriday
' 	' 	Case	"SAT",		"SATURDAY",	"VBSATURDAY":		Fld_Spec = VBA.VbDayOfWeek.vbSaturday
' 	' 	' 	============	==========	======================
' 	' 	Case Else:												GoTo NO_MATCH
' 	' 	End Select
' 		
' 	' ' The "FirstWeekOfYear" for Format().
' 	' Case FieldArgument.argWeek1
' 	' 	Select Case spec
' 	' 	' 	============	==========	======================
' 	' 	Case	"SYS",		"SYSTEM",	"VBUSESYSTEM":		Fld_Spec = VBA.VbFirstWeekOfYear.vbUseSystem
' 	' 	Case	"J1",		"JAN1",		"VBFIRSTJAN1":		Fld_Spec = VBA.VbFirstWeekOfYear.vbFirstJan1
' 	' 	Case	"4D",		"FOURDAYS",	"VBFIRSTFOURDAYS":	Fld_Spec = VBA.VbFirstWeekOfYear.vbFirstFourDays
' 	' 	Case	"FW",		"FULLWEEK",	"VBFIRSTFULLWEEK":	Fld_Spec = VBA.VbFirstWeekOfYear.vbFirstFullWeek
' 	' 	' 	============	==========	======================
' 	' 	Case Else:												GoTo NO_MATCH
' 	' 	End Select
' 		
' 	' .
' 	Case FieldArgument.argFormat
' 		Select Case spec
' 		' 	============	==========	======================
' 	' ' 	Case	"?",		"UNKNOWN",	"[_Unknown]":		Fld_Spec = FormatMode.[_Unknown]
' 		Case	"XL",		"EXCEL",	"FMTEXCELTEXT":		Fld_Spec = FormatMode.fmtExcelText
' 		Case	"VB",		"VBA",		"FMTVBFORMAT":		Fld_Spec = FormatMode.fmtVbFormat
' 		' 	============	==========	======================
' 		Case Else:												GoTo NO_MATCH
' 	' 	Case Else:							Fld_Spec = FormatMode.[_Unknown]:	GoTo NO_MATCH
' 		End Select
' 	End Select
' 	
' 	exists = True
' 	Exit Function
' 	
' NO_MATCH:
' 	exists = False
' End Function


' Close an expression and record its information.
Private Function Expr_Close(ByRef expr As ParserExpression, _
	ByRef format As String _
) As ParsingStatus
	' Record the syntax...
	If expr.Start <= expr.Stop Then
		Dim exprLen As Long: exprLen = expr.Stop - expr.Start + 1
		expr.Syntax = VBA.Mid$(format, expr.Start, exprLen)
		
	' ...or clear invalid information.
	Else
		Expr_Reset expr
	End If
	
	' This should always work.
	Expr_Close = ParsingStatus.stsSuccess
End Function


' ' Close an element and record its information.
' Private Sub Elm_Close(ByRef elm As ParserElement, _
' 	ByRef format As String _
' )
' 	' Close the expression.
' 	Expr_Close elm.Expression, format := format
' End Sub


' ' Close an element and record its information.
' Private Function Elm_Close(ByRef elm As ParserElement, _
' 	ByRef format As String, _
' 	ByRef nDfu As Long, _
' 	ByRef idxEsc As Boolean _
' ) As ParsingStatus
' 	Dim status As ParsingStatus
' 	Elm_Close = ParsingStatus.stsSuccess
' 	
' 	' Record any error when closing its expression.
' 	status = Expr_Close(elm.Expression, format := format)
' 	If Elm_Close = ParsingStatus.stsSuccess Then Elm_Close = status
' 	
' 	' Record any error when closing its extended (sub)element.
' 	Select Case elm.Kind
' 	Case ElementKind.elmField
' 		status = Fld_Close(elm.Field, format := format, nDfu := nDfu, idxEsc := idxEsc)
' 	Case Else
' 		status = ParsingStatus.stsSuccess
' 	End Select
' 	
' 	If Elm_Close = ParsingStatus.stsSuccess Then Elm_Close = status
' End Function


' Close a field (sub)element and record its information...
Private Function Fld_Close(ByRef fld As ParserField, _
	ByRef format As String, _
	ByRef nDfu As Long, _
	ByRef idxEsc As Boolean _
) As ParsingStatus
	Dim status As ParsingStatus
	Fld_Close = ParsingStatus.stsSuccess
	
	' Record any error when closing its index...
	status = Idx_Close(fld.Index, format := format, nDfu := nDfu, idxEsc := idxEsc)
	If Fld_Close = ParsingStatus.stsSuccess Then Fld_Close = status
	
	' ...and its format.
	status = Fmt_Close(fld.Format, format := format)
	If Fld_Close = ParsingStatus.stsSuccess Then Fld_Close = status
End Function


' ...along with its index (sub)element.
Private Function Idx_Close(ByRef idx As ParserIndex, _
	ByRef format As String, _
	ByRef nDfu As Long, _
	ByRef idxEsc As Boolean _
) As ParsingStatus
	Dim idxQuo As Boolean: idxQuo = False
	
	' Record the index...
	If idx.Exists And idx.Expression.Start <= idx.Expression.Stop Then
		idx.Expression.Stop = idx.Expression.Stop - 1
		Dim idxLen As Long: idxLen = idx.Expression.Stop - idx.Expression.Start + 1
		idx.Expression.Syntax = VBA.Mid$(format, idx.Expression.Start, idxLen)
		idxQuo = (nDfu = 1)
		
	' ...or clear invalid information.
	Else
		idx.Expression.Start = 0
		idx.Expression.Stop = 0
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
		
		Fld_Close = ParsingStatus.stsErrorInvalidIndex
		Exit Function
	End If
End Function


' ' ...and its format (sub)element.
' Private Function Fmt_Close(ByRef fmt As ParserFormat, _
' 	ByRef format As String _
' ) As ParsingStatus
' 	' Record the format...
' 	If fmt.Exists And fmt.Expression.Start <= fmt.Expression.Stop Then
' 		fmt.Expression.Start = fmt.Expression.Start + 1
' 		Dim fmtLen As Long: fmtLen = fmt.Expression.Stop - fmt.Expression.Start + 1
' 		fmt.Expression.Syntax = VBA.Mid$(format, fmt.Expression.Start, fmtLen)
' 		
' 	' ...or clear invalid information.
' 	Else
' 		fmt.Expression.Start = 0
' 		fmt.Expression.Stop = 0
' 	End If
' 	
' 	' This should always work.
' 	Fmt_Close = ParsingStatus.stsSuccess
' End Function



' ########################
' ## Support | Elements ##
' ########################

' Reset an expression.
Private Sub Expr_Reset(ByRef expr As ParserExpression)
	Dim reset As ParserExpression
	Let expr = reset
End Sub


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
	Let elm2.Kind   = elm1.Kind
	Let elm2.Plain	= elm1.Plain
	Fld_Clone elm1.Field, elm2.Field
End Sub


' Clone one field (sub)element into another.
Private Sub Fld_Clone(ByRef fld1 As ParserField, ByRef fld2 As ParserField)
	Let fld2.Index    = fld1.Index
' 	Let fld2.Position = fld1.Position
' 	Let fld2.Day1     = fld1.Day1
' 	Let fld2.Week1    = fld1.Week1
	Let fld2.Format   = fld1.Format
End Sub
