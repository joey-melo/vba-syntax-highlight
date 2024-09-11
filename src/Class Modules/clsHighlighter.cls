VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHighlighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Member variables
Private s_language As String
Private s_commentLine As String
Private s_commentMultilineStart As String
Private s_commentMultilineEnd As String
Private s_stringMultilineStart As String
Private s_stringMultilineEnd As String
Private r_range As Range
Private a_reserved As Variant
Private a_operators As Variant
Private a_types As Variant
Private a_builtins As Variant
Private a_literals As Variant
Private o_format As clsFormat
Private o_styles As clsStyles

' Helper type for identifying strings and comments
Private Type NextElement
    pos As Long
    char As String
    mode As String
End Type

''' Getters and Setters (Class properties)

' Format properties

Property Get Format() As clsFormat
    Set Format = o_format
End Property
Property Let Format(value As clsFormat)
    Set o_format = value
End Property
Property Get styles() As clsStyles
    Set styles = o_styles
End Property
Property Let styles(value As clsStyles)
    Set o_styles = value
End Property

' Highligter properties

Property Get Language() As String
    Language = s_language
End Property
Property Let Language(value As String)
    s_language = value
End Property
Property Get CommentLine() As String
    CommentLine = s_commentLine
End Property
Property Let CommentLine(value As String)
    s_commentLine = value
End Property
Property Get CommentMultilineStart() As String
    CommentMultilineStart = s_commentMultilineStart
End Property
Property Let CommentMultilineStart(value As String)
    s_commentMultilineStart = value
End Property
Property Get CommentMultilineEnd() As String
    CommentMultilineEnd = s_commentMultilineEnd
End Property
Property Let CommentMultilineEnd(value As String)
    s_commentMultilineEnd = value
End Property
Property Get StringMultilineStart() As String
    StringMultilineStart = s_stringMultilineStart
End Property
Property Let StringMultilineStart(value As String)
    s_stringMultilineStart = value
End Property
Property Get StringMultilineEnd() As String
    StringMultilineEnd = s_stringMultilineEnd
End Property
Property Let StringMultilineEnd(value As String)
    s_stringMultilineEnd = value
End Property
Property Get Range() As Range
    Range = r_range
End Property
Property Let Range(value As Range)
    Set r_range = ValidateRange(value)
End Property
Property Get Reserved() As Variant
    Reserved = a_reserved
End Property
Property Let Reserved(value As Variant)
    a_reserved = value
End Property
Property Get Operators() As Variant
    Operators = a_operators
End Property
Property Let Operators(value As Variant)
    a_operators = value
End Property
Property Get Types() As Variant
    Types = a_types
End Property
Property Let Types(value As Variant)
    a_types = value
End Property
Property Get Builtins() As Variant
    Builtins = a_builtins
End Property
Property Let Builtins(value As Variant)
    a_builtins = value
End Property
Property Get Literals() As Variant
    Literals = a_literals
End Property
Property Let Literals(value As Variant)
    a_literals = value
End Property
'''

' __Init__
Private Sub Class_Initialize()
    Set o_format = New clsFormat
    Set o_styles = New clsStyles
    o_format.FontName = "Consolas"
    o_format.FontSize = 9
    o_format.SpaceBefore = 10
    o_format.SpaceAfter = 10
    o_format.BorderStyle = wdLineStyleSingle
    Set r_range = ValidateRange(Selection.Range)
End Sub

' Range validator (The user must select a range that will be highlighted)
Private Function ValidateRange(rng As Range) As Range
    On Error GoTo HandleError

    ' Check if the range is valid
    If rng.start < rng.End Then
        Set ValidateRange = rng
    Else
        Err.Raise vbObjectError, "ValidateRange", "Select the text range to be highlighted"
    End If
    Exit Function

HandleError:
    MsgBox Err.Description
    End
End Function

' Syntax Highlighting Logic
Private Sub HighlightRange(posStart As Long, posEnd As Long, styles As clsFormat)
    Dim rng As Range
    Set rng = r_range.Duplicate
    
    rng.start = posStart
    rng.End = posEnd
    rng.Font.Color = styles.Color
    rng.Font.Italic = styles.Italics
End Sub

Private Sub HighlightKeywords(arrKeywords As Variant, styles As clsFormat)

    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    Dim posMax As Long
    
    Set rng = r_range
    posMax = Len(rng.text)

    ' Loop through keywords arrays
    For Each keyword In arrKeywords

        posStart = InStr(1, rng.text, keyword)
        
        ' Find all instances of the keyword in the range
        Do While posStart > 0 And posStart < posMax
            posEnd = posStart + Len(keyword)
            
             If IsWord(rng, posStart, posEnd, posMax) Then
                HighlightRange rng.start + posStart - 1, rng.start + posEnd - 1, styles

             End If
             
             posStart = InStr(posEnd + 1, rng.text, keyword)
        Loop
    Next keyword
        
End Sub
Private Function IsWord(rng As Range, posStart As Long, posEnd As Long, posMax As Long) As Boolean

    If posStart > 1 And posEnd < posMax Then
        ' Scenario 1: word is in the middle of the range
        If Not Mid(rng.text, posStart - 1, 1) Like "[0-9a-zA-Z_]" And Not Mid(rng.text, posEnd, 1) Like "[0-9a-zA-Z_]" Then IsWord = True
    ElseIf posStart = 1 And posEnd < posMax Then
        'Scenario 2: word starts in the beginning of the range but does not end at the end
        If Not Mid(rng.text, posEnd, 1) Like "[0-9a-zA-Z_]" Then IsWord = True
    ElseIf posStart > 1 And posEnd = posMax Then
        'Scenario 3: word does not start at beginning of range, but ends at the end of the range
        If Not Mid(rng.text, posStart - 1, 1) Like "[0-9a-zA-Z_]" Then IsWord = True
    Else
        'Scenario 4: word starts and ends with range (word is the sole selection)
        IsWord = True
    End If
    
End Function
Public Sub HighlightOperators()
    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    
    Set rng = r_range
    
    For Each op In a_operators
        posStart = InStr(1, rng.text, op)
        
        Do While posStart > 0 And posStart < Len(rng.text)
            posEnd = posStart + Len(op)
            HighlightRange rng.start + posStart - 1, rng.start + posEnd - 1, o_styles.Operators
            posStart = InStr(posEnd + 1, rng.text, op)
        Loop
        
    Next op
End Sub

Public Sub HighlightNumbers()
    Dim rng As Range
    Dim posEnd As Long
    Dim posStart As Long
    
    Set rng = r_range

    For Each wd In rng.Words
        If wd Like "[0-9]*" Then
            posStart = wd.start
            posEnd = wd.End

            HighlightRange posStart, posEnd, o_styles.Numbers
        End If
    Next wd
    
End Sub

Private Function GetSeparatorStartPos(rng As Range, posEnd As Long) As Long
    Dim pos As Long: pos = posEnd
    
    Do While pos > 1
        If Not Mid(rng.text, pos, 1) Like "[0-9a-zA-Z_]" Then
            GoTo assignValue
        Else
            pos = pos - 1
        End If
    Loop

assignValue:
    GetSeparatorStartPos = pos
    
End Function

Public Sub HighlightFunctionCalls()
    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    Dim posMax As Long
    
    Set rng = r_range
    posMax = Len(rng.text)
    
    posStart = InStr(1, rng.text, "(")
    
    Do While posStart > 0 And posStart < posMax
        posEnd = posStart - 1
        posStart = GetSeparatorStartPos(rng, posEnd)
        
        HighlightRange rng.start + posStart, rng.start + posEnd, o_styles.FunctionCalls
        
        posStart = InStr(posEnd + 2, rng.text, "(")
    Loop
    
End Sub
Private Function GetSeparatorEndPos(rng As Range, posStart As Long) As Long
    Dim pos As Long
    Dim regex As String
    
    pos = posStart
    regex = "[0-9a-zA-Z_]"

    Do While pos < Len(rng.text)
        If Not Mid(rng.text, pos, 1) Like regex Then
            GoTo assignValue
        Else
            pos = pos + 1
        End If
    Loop
    
assignValue:
    GetSeparatorEndPos = pos - 1
End Function
Private Function GetNextParameter(rng As Range, posStart As Long, nest As Integer) As Long
    Dim pos As Long: pos = posStart
    
    Do While pos <= Len(rng.text)
        Select Case Mid(rng.text, pos, 1)
            Case ")" ' Check for closing parenthesis
                If nest = 0 Then ' If no open parenthesis pending, it's the end
                    pos = 0
                    Exit Do
                Else ' Otherwise, decrease the count of open parentheses
                    nest = nest - 1
                End If
                
            Case "(" ' New function or nested parameter
                nest = nest + 1
                pos = pos + 1
                Exit Do
                
            Case "," ' End of parameter, move to next character
                If openParenCount = 0 Then
                    pos = pos + 1
                    Exit Do
                End If
                
            Case "=" ' Assignment, move to the next word
                pos = GetSeparatorEndPos(rng, pos + 1) + 1
                Exit Do
        End Select
        
        pos = pos + 1
    Loop
            
    If pos > 0 And pos < Len(rng.text) Then
        Do While Not Mid(rng.text, pos, 1) Like "[0-9a-zA-Z_]"
            pos = pos + 1
        Loop
    End If
        
    ' Assign the found position to return
    GetNextParameter = pos
End Function
Public Sub HighlightFunctionParameters()
    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    Dim posMax As Long
    Dim nest As Integer
    
    Set rng = r_range
    posMax = Len(rng.text)
    nest = 0 ' counter for nested functions
    
    posStart = InStr(1, rng.text, "(") + 1
    
    Do While posStart > 0 And posStart < posMax
        posEnd = GetSeparatorEndPos(rng, posStart)
        
        HighlightRange rng.start + posStart - 1, rng.start + posEnd, o_styles.FunctionParameters
        
        posStart = GetNextParameter(rng, posEnd + 1, nest)

    Loop

End Sub
Private Function IsEscaped(rng As Range, pos As Long) As Boolean
    IsEscaped = False
    If pos > 1 Then
        IsEscaped = Mid(rng.text, pos - 1, 1) = "\"
    End If
End Function
Public Sub HighlightStringsAndComments()
    ' This sub will categorize the range into comment, string, or free
    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    Dim posMax As Long
    Dim element As NextElement
    
    Set rng = r_range
    posMax = Len(rng.text)

    posStart = 1
    Do While posStart > 0 And posStart < posMax
        element = NextCommentOrString(posStart)
        posStart = element.pos
        posEnd = InStr(posStart + 1, rng, element.char) + Len(element.char) - 1
        While IsEscaped(rng, posEnd)
            posEnd = InStr(posEnd + 1, rng, element.char) + Len(element.char) - 1
        Wend
        
        ' If it's an open element, highlight everything
        If posEnd = 0 Or posEnd = Len(element.char) - 1 Then posEnd = posMax
        If posStart > 0 And posStart < posEnd Then
            If element.mode = "comment" Then
                HighlightRange r_range.start - 1 + posStart, r_range.start + posEnd, o_styles.Comments
            Else
                HighlightRange r_range.start - 1 + posStart, r_range.start + posEnd, o_styles.Strings
            End If
        End If
        posStart = posEnd + 1
    Loop

End Sub

Private Function NextCommentOrString(posStart As Long) As NextElement
    Dim result As NextElement
    Dim mode As String
    Dim posComment As Long
    Dim posMultilineComment As Long
    Dim posStringMultiline As Long
    Dim posString1 As Long
    Dim posString2 As Long
    
    If Len(s_commentLine) > 0 Then posComment = InStr(posStart, r_range.text, s_commentLine)
    If Len(s_commentMultilineStart) > 0 Then posMultilineComment = InStr(posStart, r_range.text, s_commentMultilineStart)
    If Len(s_stringMultilineStart) > 0 Then posStringMultiline = InStr(posStart, r_range.text, s_stringMultilineStart)
    posString1 = InStr(posStart, r_range.text, Chr(34)) ' double quotes (")
    posString2 = InStr(posStart, r_range.text, Chr(39)) ' single quotes (')
    
    ' Basically a MIN() function that is not supported by VBA
    posStart = Len(r_range)
    If posComment > 0 And posComment < posStart And Not IsEscaped(r_range, posComment) Then
        posStart = posComment
        charEnd = Chr(13)
        mode = "comment"
    End If
    If posMultilineComment > 0 And posMultilineComment < posStart And Not IsEscaped(r_range, posMultilineComment) Then
        posStart = posMultilineComment
        charEnd = s_commentMultilineEnd
        mode = "comment"
    End If
    If posStringMultiline > 0 And posStringMultiline < posStart And Not IsEscaped(r_range, posStringMultiline) Then
        posStart = posStringMultiline
        charEnd = s_stringMultilineEnd
        mode = "string"
    End If
    If posString1 > 0 And posString1 < posStart And Not IsEscaped(r_range, posString1) Then
        posStart = posString1
        charEnd = Chr(34) ' (")
        mode = "string"
    End If
    If posString2 > 0 And posString2 < posStart And Not IsEscaped(r_range, posString2) Then
        posStart = posString2
        charEnd = Chr(39) ' (')
        mode = "string"
    End If
    
    result.pos = posStart
    result.char = charEnd
    result.mode = mode
    NextCommentOrString = result
End Function

Public Sub HighlightCmdArguments()

    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    Dim posMax As Long
    Dim posSpace As Long
    Dim posParagraph As Long
    Dim regex As String
    
    Set rng = r_range.Duplicate
    posMax = Len(rng.text)
    regex = "[ " & Chr(13) & "]" ' space or paragraph mark (13)
    
    posStart = InStr(1, rng.text, "-")
    
    Do While posStart > 0 And posStart < posMax
        posEnd = posStart + 1
        
        ' Find next space or paragraph mark
        Do While posEnd < posMax
            If Mid(rng.text, posEnd, 1) Like regex Then Exit Do
            posEnd = posEnd + 1
        Loop

        If IsWord(rng, posStart, posEnd, posMax) Then
            HighlightRange rng.start + posStart - 1, rng.start + posEnd, o_styles.CmdArguments
        End If
        
        posStart = InStr(posEnd + 1, rng.text, "-")
    Loop

End Sub

Public Sub HighlightVariables()
    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    Dim posMax As Long
    
    Set rng = r_range
    posMax = Len(rng.text)
    
    posStart = InStr(1, rng.text, "$")
    
    Do While posStart > 0 And posStart < posMax
        posEnd = GetSeparatorEndPos(rng, posStart + 1)
        
        HighlightRange rng.start + posStart - 1, rng.start + posEnd, o_styles.Variables
        
        posStart = InStr(posEnd + 1, rng.text, "$")
    Loop
End Sub
Public Sub HighlightTags()
    Dim rng As Range
    Dim posStart As Long
    Dim posEnd As Long
    Dim posMax As Long
    
    Set rng = r_range
    posMax = Len(rng.text)
    
    posStart = InStr(1, rng.text, "<")
    
    Do While posStart > 0 And posStart < posMax
        If Mid(rng.text, posStart + 1, 1) = "/" Then posStart = posStart + 1
        
        posEnd = GetSeparatorEndPos(rng, posStart + 1)
        
        If posEnd > posStart Then HighlightRange rng.start + posStart, rng.start + posEnd, o_styles.Tags
        
        posStart = InStr(posEnd + 1, rng.text, "<")
    Loop
    
End Sub

Public Sub PrepareRange()
    o_format.CleanRange r_range ' Remove previous formatting and replace bad characters
    o_format.ApplyBorders r_range
    o_format.ApplySpacing r_range, o_format.SpaceBefore, o_format.SpaceAfter
    o_format.SetFont r_range, o_format.FontName, o_format.FontSize
    o_format.DisableProofing r_range ' Disable grammar and spell check
End Sub
Public Sub HighlightLiterals()
    HighlightKeywords a_literals, o_styles.KeywordsLiterals
End Sub
Public Sub HighlightReserved()
    HighlightKeywords a_reserved, o_styles.KeywordsReserved
End Sub
Public Sub HighlightBuiltins()
    HighlightKeywords a_builtins, o_styles.KeywordsBuiltins
End Sub
Public Sub HighlightTypes()
    HighlightKeywords a_types, o_styles.KeywordsTypes
End Sub

Private Sub ExecuteHighlight(ByVal func As String)
    Dim startTime As Single
    Dim elapsedTime As Single
    
    startTime = Timer
    CallByName Me, func, VbMethod
    elapsedTime = Timer - startTime
    Debug.Print func & " completed in: " & elapsedTime & " seconds"
End Sub

Public Sub Highlight()
    Dim dbgStart As Single
    Dim dbgEnd As Single
    
    Dim arrDefault As Variant
    Dim arrCode As Variant
    Dim arrShell As Variant
    Dim arrMarkup As Variant
    Dim arrExecute As Variant
    
    ' Will be executed regardless of language style
    arrDefault = Array( _
        "HighlightFunctionCalls", _
        "HighlightLiterals", _
        "HighlightReserved", _
        "HighlightBuiltins", _
        "HighlightTypes", _
        "HighlightStringsAndComments" _
    )
    
    ' Will execute if language style is Code
    arrCode = Array( _
        "HighlightFunctionParameters", _
        "HighlightOperators", _
        "HighlightNumbers" _
    )
    
    ' Will execute if language style is Shell
    arrShell = Array( _
        "HighlightCmdArguments", _
        "HighlightVariables" _
    )
        
    ' Will execute if language style is Markup
    arrMarkup = Array( _
        "HighlightTags" _
    )
    
    dbgStart = Timer

    Select Case s_language
        Case "Shell"
            arrExecute = arrShell

        Case "Code"
            arrExecute = arrCode
            
        Case "Markup"
            arrExecute = arrMarkup
    End Select
    
    ' Concatenate arrays
    arrExecute = Split(Join(arrExecute, ",") & "," & Join(arrDefault, ","), ",")
    
    For Each func In arrExecute
        ExecuteHighlight func
    Next func

    dbgEnd = Timer - dbgStart
    Debug.Print "Total execution time: " & dbgEnd & " seconds" & vbNewLine
End Sub




