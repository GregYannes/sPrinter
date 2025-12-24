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
	stsSuccess               =    0	' Report success.
	stsError                 = 1000	' Report a general syntax error.
	stsErrorHangingEscape    = 1001	' Report a hanging escape...
	stsErrorUnenclosedField  = 1002	' ...or an incomplete field...
	stsErrorUnenclosedQuote  = 1003	' ...or an incomplete quote...
	stsErrorInvalidIndex     = 1004	' ...or an index that is not an integer...
' 	stsErrorInvalidSpecifier = 1005	' ...or an unrecognized specifier.
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
' 	argMode					' The engine used for formatting.
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
	Index As Variant		' The index to extract the value...
' 	Position As PositionKind	' ...and how to interpret that index.
' 	Mode As FormatMode		' The engine used for formatting...
' 	Day1 As VBA.VbDayOfWeek		' ...and the conventions used by Format() for weekdays...
' 	Week1 As VBA.VbFirstWeekOfYear	' ...and for calendar weeks.
	Format As String		' The formatting code applied to the value.
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
	' ' ###############################
	' Dim eIdx As Long: eIdx = base - 1
	' ' ###############################
	Dim eIdx As Long: eIdx = base
	Dim e As ParserElement
	
	' ...and the current (field) argument...
	Dim args(FieldArgument.[_First] To FieldArgument.[_Last]) As ParserExpression
	Dim argIdx As FieldArgument: argIdx = FieldArgument.[_None]
	Dim arg As ParserExpression
	Dim idxDfu As ParserExpression
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
		
	' Revisit the character from scratch.
	SAME_CHAR:
		' Interpret this character in context.
		Select Case e.Kind
		
		
		
		' ##############
		' ## Inactive ##
		' ##############
		
		Case ElementKind.[_Unknown]
			' ' #################################################
			' ' Save the last element...
			' If eIdx >= base Then Elm_Clone e, elements(eIdx)
			' 
			' ' ...and advance to the next.
			' eIdx = eIdx + 1
			' 
			' ' Reset the global trackers:
			' Expr_Reset expression	' Clear the expression.
			' Elm_Reset e		' Clear the element contents.
			' 
			' ' ' Enter parsing.
			' ' depth = depth + 1
			' 
			' ' Locate the element.
			' expression.Start = charIndex
			' expression.Stop = expression.Start
			' ' #################################################
			
			
			Select Case char
			
			' Parse into a field...
			Case openField
				' ' ###################################################################
				' ' Reset the global (field) trackers:
				' Erase args				' Clear the array of arguments.
				' argIdx = FieldArgument.[_None]	' Reset to no arguments.
				' nDfu = 0				' Reset to no quotations.
				' idxEsc = False			' Reset to unescaped.
				' ' ###################################################################
				
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
				expression.Stop = expression.Start - 1
				
				' Revisit this character in plaintext.
				' expression.Stop = expression.Stop - 1
				GoTo SAME_CHAR
			End Select
		End If
		
		
		
		' ################
		' ## Plain Text ##
		' ################
		
		Case ElementKind.elmPlain
			' Escape a literal character...
			If Enum_Has(ParsingDefusal.dfuEscape) Then
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
			Case Else
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
					' ' #########################################
					' ' Exit parsing so we can reenter the field.
					' depth = depth - 1
					' ' #########################################
					
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
			End Select
			
			
			
		' ###########
		' ## Field ##
		' ###########
		
		Case ElementKind.elmField
			
			' ' ################################################################################################################################################
			' ' Handle the field contents.
			' If depth = 1 Then
			' 	' ' #################################################################
			' 	' ' Save the last argument.
			' 	' If argIdx > FieldArgument.[_None] Then Expr_Clone arg, args(argIdx)
			' 	' 
			' 	' ' Reset the local (argument) trackers.
			' 	' Expr_Reset arg	' Clear the argument contents.
			' 	' ' #################################################################
			' 	
			' 	Select Case char
			' 	
			' 	' Parse out of the field...
			' 	Case closeField
			' 		' Extend the (element) location.
			' 		expression.Stop = expression.Stop + 1
			' 		
			' 		' Exit parsing so we can reenter the next element.
			' 		depth = depth - 1
			' 		
			' 		' Save this (field) element to the array...
			' 		endStatus = Fld_Close(e.Field, format := format, expression := expression, args := args, argIdx := argIdx, idxDfu := idxDfu, idxEsc := idxEsc)
			' 		If endStatus = ParsingStatus.psSuccess Then
			' 			Elm_Clone e, elements(eIdx)
			' 			
			' 		' ...unless there is a parsing error.
			' 		Else
			' 			GoTo EXIT_LOOP
			' 		End If
			' 		
			' 		' Advance to the next element.
			' 		eIdx = eIdx + 1
			' 		
			' 		' Reset the global trackers.
			' 		Expr_Reset expression
			' 		Elm_Reset e
			' 		Erase args
			' 		argIdx = FieldArgument.[_None]
			' 		nDfu = 0
			' 		idxEsc = False
			' 		
			' 		' Advance to the next character.
			' 		GoTo NEXT_CHAR
			' 		
			' 	' ...or parse into the next argument...
			' 	Case separator
			' 		' Extend the (element) location.
			' 		expression.Stop = expression.Stop + 1
			' 		
			' 		' .
			' 		
			' 		
			' 		' Prepare the argument.
			' 		arg.Stop = charIndex
			' 		arg.Start = arg.Stop + 1
			' 		
			' 		' Advance to the next argument.
			' 		argIdx = argIdx + 1
			' 		
			' 		' Advance to the next character.
			' 		GoTo NEXT_CHAR
			' 		
			' 	' ...or parse into the current argument.
			' 	Case Else
			' 		' Ensure arguments are counted and prepared.
			' 		If argIdx = FieldArgument.[_None] Then
			' 			argIdx = argIdx + 1
			' 			arg.Start = charIndex
			' 			arg.Stop = arg.Start
			' 		End If
			' 		
			' 		' Enter the argument.
			' 		depth = depth + 1
			' 		
			' 		' Revisit this character in the argument.
			' 		arg.Stop = arg.Stop - 1
			' 		GoTo SAME_CHAR
			' 	End Select
			' End If
			' ' ################################################################################################################################################
			
			
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
					If depth = 1 And argIndex = FieldArguments.argIndex Then idxDfu.Stop = charIndex
					
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
						If argIdx = FieldArguments.argIndex Then idxDfu.Stop = charIndex
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
			Case Else
				Select Case char
				
				' Escape the next character...
				Case escape
					' Confirm the argument.
					If argIdx = FieldArguments.[_None] Then argIdx = argIdx + 1
					
					' Activate escaping.
					dfu = dfu + ParsingDefusal.dfuEscape
					
					' Note any escaping in the index argument.
					If depth = 1 And argIdx = FieldArguments.argIndex Then idxEsc = True
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument.
					arg.Stop = arg.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or quote the next characters...
				Case openQuote
					' Confirm the argument.
					If argIdx = FieldArguments.[_None] Then argIdx = argIdx + 1
					
					' Activate quoting.
					dfu = dfu + ParsingDefusal.dfuQuote
					
					' Locate any quoting in the index argument.
					If depth = 1 And argIndex = FieldArguments.argIndex Then idxDfu.Start = charIndex
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument.
					arg.Stop = arg.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or nest the next syntax...
				Case openField
					' Confirm the argument.
					If argIdx = FieldArguments.[_None] Then argIdx = argIdx + 1
					
					' Nest deeper into the field.
					depth = depth + 1
					
					' Activate nesting.
					dfu = dfu + ParsingDefusal.dfuNest
					
					' Locate any nesting in the index argument.
					If depth = 1 And argIndex = FieldArguments.argIndex Then idxDfu.Start = charIndex
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument.
					arg.Stop = arg.Stop + 1
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or parse out of the field...
				Case closeField
					' Unnest out of the field.
					depth = depth - 1
					
					' Extend the location of this field.
					expression.Stop = expression.Stop + 1
					
					' Save (and reset) any argument to the array...
					If argIdx > FieldArguments.[_None] Then
						Expr_Clone arg, args(argIdx)
						Expr_Reset arg
					End If
					
					' ...along with the (field) element.
					endStatus = Fld_Close(e.Field, format := format, expression := expression, args := args, argIdx := argIdx, idxDfu := idxDfu, idxEsc := idxEsc)
					Elm_Clone e, elements(eIdx)
					Elm_Reset e
					
					' Short-circuit for errors.
					If endStatus <> ParsingStatus.stsSuccess Then GoTo EXIT_LOOP
					
					' Advance to the next element.
					eIdx = eIdx + 1
					
					' Reset the global trackers.
					Expr_Reset expression
					Erase args
					argIdx = FieldArgument.[_None]
					Expr_Reset idxDfu
					idxEsc = False
					
					' Advance to the next character.
					GoTo NEXT_CHAR
					
				' ...or parse to the next argument...
				Case separator
					' Confirm the argument.
					If argIdx = FieldArguments.[_None] Then argIdx = argIdx + 1
					
					' Extend the location of this field.
					expression.Stop = expression.Stop + 1
					
					' Begin a new argument...
					If argIdx < FieldArguments.[_Last] Then
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
					' Confirm the argument.
					If argIdx = FieldArguments.[_None] Then argIdx = argIdx + 1
					
					' Extend the location of this field...
					expression.Stop = expression.Stop + 1
					
					' ...and of this argument...
					arg.Stop = arg.Stop + 1
					
					' ...along with its contents.
					arg.Syntax = arg.Syntax & char
					
					' Advance to the next character.
					GoTo NEXT_CHAR
				End Select
			End Select
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
	
	' Resize to the elements we actually parsed.
	If e.Kind = ElementKind.[_Unknown] Then
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
			expression := expression, _
			args := args, _
			argIdx := argIdx, _
			idxDfu := idxDfu, _
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

' ' Get a missing (Variant) argument.
' Private Function Missing(Optional ByRef x As Variant) As Variant
' 	Let Missing = x
' End Function



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
' 	Optional ByRef idxDfu As ParserExpression, _
' 	Optional ByRef idxEsc As Boolean, _
' 	Optional ByRef endStatus As ParsingStatus _
' )
' 	dfu = ParsingDefusal.[_Off]
' 	depth = 0
' 	eIdx = 0
' 	Elm_Reset e
' 	char = VBA.vbNullString
' 	Expr_Reset idxDfu
' 	idxEsc = False
' 	endStatus = ParsingStatus.stsSuccess
' End Sub



' #################################
' ## Support | Parsing | Closure ##
' #################################

' Close an expression and record its information.
Private Sub Expr_Close(ByRef expr As ParserExpression, _
	ByRef format As String _
)
	' Record the syntax...
	If expr.Start > 0 And expr.Start <= expr.Stop Then
		Dim exprLen As Long: exprLen = expr.Stop - expr.Start + 1
		expr.Syntax = VBA.Mid$(format, expr.Start, exprLen)
		
	' ...or clear invalid information.
	Else
		expr.Syntax = VBA.vbNullString
		expr.Stop = 0
	End If
End Sub


' Close a field (sub)element and record its information...
Private Function Fld_Close(ByRef fld As ParserField, _
	ByRef format As String, _
	ByRef expression As ParserExpression, _
	ByRef args() As ParserExpression, _
	ByRef argIdx As Long, _
	ByRef idxDfu As ParserExpression, _
	ByRef idxEsc As Boolean _
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
	Dim iTo As Long: iTo = argIdx - 1
	
	Dim i As Long
	For i = FieldArgument.[_First] To iTo
		' Extract the argument by position.
		arg = args(i)
		' Expr_Clone args(i), arg
		
		
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
			
		' ' Process a specifier argument.
		' Case FieldArgument.argPosition, FieldArgument.argMode, FieldArgument.argDay1, FieldArgument.argWeek1
		' 	Fld_Close = Fld_CloseSpecifier(fld, _
		' 		arg := i, _
		' 		spec := arg, _
		' 		format := format, _
		' 		expression := expression _
		' 	)
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
	ByRef idxEsc As Boolean _
) As ParsingStatus
	' Define fallback for missing argument.
	Dim noIdx As Variant  ' noIdx = Missing()
	
	' Save the defused syntax...
	Dim dfuSyntax As String: dfuSyntax = idx.Syntax
	
	' ...before recording the original syntax.
	Expr_Close idx, format := format
	
	' Short-circuit for a missing index.
	If idx.Syntax = VBA.vbNullString Then
		Let fld.Index = noIdx
		Fld_CloseIndex = ParsingStatus.stsSuccess
		Exit Function
	End If
	
	' Check if the index is encapsulated in a single quotation ("...") or nesting ({...}).
	Dim idxCap As Boolean: idxCap = (idxDfu.Start = idx.Start And idxDfu.Stop = idx.Stop)
	
	' Interpret as an (encapsulated) key...
	If idxCap Or idxEsc Then
		Let fld.Index = VBA.CStr(dfuSyntax)
		
	' ...or an integral index.
	Else
		On Error GoTo IDX_ERROR
		Let fld.Index = VBA.CLng(idx.Syntax)
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


' ' ...and any of its specifier arguments...
' Private Function Fld_CloseSpecifier(ByRef fld As ParserField, _
' 	ByRef arg As FieldArgument, _
' 	ByRef spec As ParserExpression, _
' 	ByRef format As String, _
' 	ByRef expression As ParserExpression _
' ) As ParsingStatus
' 	' Define fallback for missing argument.
' 	Dim noArg As Variant: noArg = -1
' 	
' 	' Record and extract the original syntax.
' 	Expr_Close spec, format := format
' 	Dim specStx As String: specStx = spec.Syntax
' 	
' 	' Clean that syntax...
' 	specStx = VBA.Trim(specStx)
' 	
' 	' ...and short-circuit for a missing argument altogether.
' 	If specStx = VBA.vbNullString Then
' 		Let Fld_Arg(fld, arg) = noArg
' 		Fld_CloseSpecifier = ParsingStatus.stsSuccess
' 		Exit Function
' 	End If
' 	
' 	' Look up the specifier that matches this syntax...
' 	Dim val As Long, exists As Boolean
' 	val = Arg_Specifier(arg, spec := specStx, exists := exists)
' 	
' 	' ...and short-circuit for no match.
' 	If Not exists Then GoTo SPEC_ERROR
' 	
' 	' Assign any valid specifier to the argument...
' 	Let Fld_Arg(fld, arg) = val
' 	
' 	' ...and report success.
' 	Expr_Reset expression
' 	Fld_CloseSpecifier = ParsingStatus.stsSuccess
' 	Exit Function
' 	
' 	
' ' Report the error for an invalid specifier.
' SPEC_ERROR:
' 	Expr_Clone spec, expression
' 	Fld_CloseSpecifier = ParsingStatus.stsErrorInvalidSpecifier
' End Function


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



' #################################
' ## Support | Elements | Fields ##
' #################################

' Clone one field (sub)element into another.
Private Sub Fld_Clone(ByRef fld1 As ParserField, ByRef fld2 As ParserField)
	Let fld2.Index    = fld1.Index
' 	Let fld2.Position = fld1.Position
' 	Let fld2.Day1     = fld1.Day1
' 	Let fld2.Week1    = fld1.Week1
	Let fld2.Format   = fld1.Format
End Sub



' #############################################
' ## Support | Elements | Fields | Arguments ##
' #############################################

' Set an argument for a field.
Private Property Let Fld_Arg(ByRef fld As ParserField, _
	ByRef arg As FieldArgument, _
	ByRef val As Variant _
)
	Select Case arg
	Case FieldArgument.argIndex:	Let fld.Index    = val
	Case FieldArgument.argPosition:	Let fld.Position = val
	Case FieldArgument.argMode:	Let fld.Mode     = val
	Case FieldArgument.argDay1:	Let fld.Day1     = val
	Case FieldArgument.argWeek1:	Let fld.Week1    = val
	Case FieldArgument.argFormat:	Let fld.Format   = val
	End Select
End Property


' ' Interpret specifiers for field arguments.
' Private Function Arg_Specifier(ByRef arg As FieldArgument, _
' 	ByVal spec As String, _
' 	Optional ByRef exists As Boolean _
' ) As Long
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
' 	' ' 	Case	"?", 		"UNKNOWN",	"[_UNKNOWN]":		Arg_Specifier = PositionKind.[_Unknown]
' 	' 	Case	"ABS",		"ABSOLUTE",	"POSABSOLUTE":		Arg_Specifier = PositionKind.posAbsolute
' 	' 	Case	"REL",		"RELATIVE",	"POSRELATIVE":		Arg_Specifier = PositionKind.posRelative
' 	' 	' 	============	==========	======================
' 	' 	Case Else: GoTo NO_MATCH
' 	' 	End Select
' 		
' 	' ' The engine used for formatting.
' 	' Case FieldArgument.argMode
' 	' 	Select Case spec
' 	' 	' 	============	==========	======================
' 	' ' 	Case	"?",		"UNKNOWN",	"[_Unknown]":		Arg_Specifier = FormatMode.[_Unknown]
' 	' 	Case	"XL",		"EXCEL",	"FMTEXCELTEXT":		Arg_Specifier = FormatMode.fmtExcelText
' 	' 	Case	"VB",		"VBA",		"FMTVBFORMAT":		Arg_Specifier = FormatMode.fmtVbFormat
' 	' 	' 	============	==========	======================
' 	' 	Case Else: GoTo NO_MATCH
' 	' 	End Select
' 		
' 	' ' The "FirstDayOfWeek" for Format().
' 	' Case FieldArgument.argDay1
' 	' 	Select Case spec
' 	' 	' 	============	==========	======================
' 	' 	Case	"SYS",		"SYSTEM",	"VBUSESYSTEMDAYOFWEEK":	Arg_Specifier = VBA.VbDayOfWeek.vbUseSystemDayOfWeek
' 	' 	Case	"SUN",		"SUNDAY",	"VBSUNDAY":		Arg_Specifier = VBA.VbDayOfWeek.vbSunday
' 	' 	Case	"MON",		"MONDAY",	"VBMONDAY":		Arg_Specifier = VBA.VbDayOfWeek.vbMonday
' 	' 	Case	"TUE",		"TUESDAY",	"VBTUESDAY":		Arg_Specifier = VBA.VbDayOfWeek.vbTuesday
' 	' 	Case	"WED",		"WEDNESDAY",	"VBWEDNESDAY":		Arg_Specifier = VBA.VbDayOfWeek.vbWednesday
' 	' 	Case	"THU",		"THURSDAY",	"VBTHURSDAY":		Arg_Specifier = VBA.VbDayOfWeek.vbThursday
' 	' 	Case	"FRI",		"FRIDAY",	"VBFRIDAY":		Arg_Specifier = VBA.VbDayOfWeek.vbFriday
' 	' 	Case	"SAT",		"SATURDAY",	"VBSATURDAY":		Arg_Specifier = VBA.VbDayOfWeek.vbSaturday
' 	' 	' 	============	==========	======================
' 	' 	Case Else: GoTo NO_MATCH
' 	' 	End Select
' 		
' 	' ' The "FirstWeekOfYear" for Format().
' 	' Case FieldArgument.argWeek1
' 	' 	Select Case spec
' 	' 	' 	============	==========	======================
' 	' 	Case	"SYS",		"SYSTEM",	"VBUSESYSTEM":		Arg_Specifier = VBA.VbFirstWeekOfYear.vbUseSystem
' 	' 	Case	"J1",		"JAN1",		"VBFIRSTJAN1":		Arg_Specifier = VBA.VbFirstWeekOfYear.vbFirstJan1
' 	' 	Case	"4D",		"FOURDAYS",	"VBFIRSTFOURDAYS":	Arg_Specifier = VBA.VbFirstWeekOfYear.vbFirstFourDays
' 	' 	Case	"FW",		"FULLWEEK",	"VBFIRSTFULLWEEK":	Arg_Specifier = VBA.VbFirstWeekOfYear.vbFirstFullWeek
' 	' 	' 	============	==========	======================
' 	' 	Case Else: GoTo NO_MATCH
' 	' 	End Select
' 	End Select
' 	
' 	exists = True
' 	Exit Function
' 	
' NO_MATCH:
' 	exists = False
' End Function
