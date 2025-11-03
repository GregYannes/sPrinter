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
Public Enum sFormatMode
	fmVbFormat	' The Format() function in VBA.
	fmXlText	' The Text() function in Excel.
End Enum


' Outcomes for parsing.
Public Enum sParseStatus
	psSuccess = 0			' Report success.
	psError = 1000			' Report a general syntax error.
	psErrorHangingEscape = 1001	' Report a hanging escape...
	psErrorUnclosedField = 1002	' ...or an incomplete field...
	psErrorUnclosedQuote = 1003	' ...or an incomplete quote...
	psErrorNonintegralIndex = 1004	' ...or an index that is not an integer.
End Enum


' Kinds of elements for parsing.
Public Enum sParseKind
	[_Unknown]	' Uninitialized.
	pkPlain		' Plain text which is displayed as is.
	pkField		' Field that is formatted and embedded.
End Enum


' Modes for parsing.
Private Enum sParseMode
	[_Off]		' Inactive.
	pmPlain		' Plain text.
	pmField		' An embedded field...
	pmFieldIndex	' ...its index...
	pmFieldFormat	' ...and its format.
End Enum



' ###########
' ## Types ##
' ###########

' Elements into which formats are parsed.
Public Type sParseElement
	Kind As sParseKind
	Text As String
	HasIndex As Boolean
	Index As String
	RawIndex As String
	IndexIsKey As String
	HasFormat As Boolean
	Format As String
End Type



