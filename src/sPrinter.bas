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
Private Const SYM_SEP As String = ":"			' Separate specifiers in a field.



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
' '	=============	==========================	  =============		=========
' '	Label		Code				  Name			Character
' '	=============	==========================	  =============		=========
' 	symEscape     =	                        92	' Backslash		\
' 	symOpenField  =	                       173	' Opening brace		{
' 	symCloseField =	                       175	' Closing brace		}
' 	symOpenQuote  =	                        34	' Double quotes		"
' 	symCloseQuote =	ParsingSymbol.symOpenQuote	' Double quotes		"
' 	symSeparator  =	                        58	' Colon			:
' End Enum


' Outcomes of parsing.
Public Enum ParsingStatus
	stsSuccess                =    0	' Report success.
	stsError                  = 1000	' Report a general syntax error.
	stsErrorHangingEscape     = 1001	' Report a hanging escape...
	stsErrorUnenclosedQuote   = 1002	' ...or an incomplete quote...
	stsErrorImbalancedNesting = 1003	' ...or an imbalanced nesting...
	stsErrorInvalidIndex      = 1004	' ...or an index that is not an integer.
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
	Index As Variant		' The index to extract the value.
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

' ###################
' ## API | Parsing ##
' ###################

' Parse a format string (without guardrails) and record granular details.
Public Sub Parse0( _
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


' ' Trim all whitespace characters from the end(s) of a string.
' Public Function Txt_Trim(ByVal txt As String, _
' 	Optional ByRef nLeft As Long, _
' 	Optional ByRef nRight As Long _
' ) As String
' 	' Count the original characters.
' 	Dim nTxt As Long: nTxt = VBA.Len(txt)
' 	
' 	' Remove all nonprinting characters...
' 	Dim cln As String: cln = txt
' 	cln = Application.WorksheetFunction.Clean$(cln)
' 	
' 	' ...and surrounding whitespace.
' 	cln = VBA.Trim$(cln)
' 	
' 	' Short-circuit for a blank result.
' 	If cln = VBA.vbNullString Then
' 		Txt_Trim = VBA.vbNullString
' 		nLeft = 0
' 		nRight = nTxt
' 		Exit Function
' 	End If
' 	
' 	' Identify the bookend (printing) characters...
' 	Dim lChr As String: lChr = VBA.Left$(cln, 1)
' 	Dim rChr As String: rChr = VBA.Right$(cln, 1)
' 	
' 	' ...and locate them in the original string.
' 	Dim lPos As Long: lPos = VBA.InStr( _
' 		String1 := txt, _
' 		String2 := lChr, _
' 		Start := 1, _
' 		Compare := VBA.VbCompareMethod.vbBinaryCompare _
' 	)
' 	Dim rPos As Long: rPos = VBA.InStrRev( _
' 		StringCheck := txt, _
' 		StringMatch := rChr, _
' 		Start := -1, _
' 		Compare := VBA.VbCompareMethod.vbBinaryCompare _
' 	)
' 	
' 	' Count the offset...
' 	nLeft = lPos - 1
' 	nRight = nTxt - rPos
' 	
' 	' ...and return the substring between those bookends.
' 	Dim nChr As Long: nChr = rPos - lPos + 1
' 	Txt_Trim = VBA.Mid$(txt, lPos, nChr)
' End Function



' #############
' ## Support ##
' #############

' ##########################
' ## Support | Validation ##
' ##########################

' Convert any input (text or code) into a valid symbol.
Private Function AsSym(ByRef x As Variant) As String
	' Extract the first character from a string...
	If VBA.VarType(x) = VBA.VbVarType.vbString Then
		AsSym = VBA.Left$(x, 1)
		
	' ...or convert a code into its character.
	Else
		#If Mac Then
			AsSym = VBA.Chr(x)
		#Else
			AsSym = VBA.ChrW(x)
		#End If
	End If
	
	' Ensure the symbol is not whitespace...
	AsSym = Application.WorksheetFunction.Clean$(AsSym)
	AsSym = VBA.Trim$(AsSym)
	If AsSym = VBA.vbNullString Then GoTo BLANK_ERROR
	
	' ...and return the result.
	Exit Function
	
	
' Throw an error for whitespace.
BLANK_ERROR:
	Err.Raise _
		Number := 5, _
		Description := "Whitespace may not be used as a formatting symbol."
End Function


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
	Err.Raise _
		Number := 5, _
		Description := "The same formatting symbol may not used twice: " & sym
End Sub



' #######################
' ## Support | Parsing ##
' #######################

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
