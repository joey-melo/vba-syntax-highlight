Attribute VB_Name = "Highlight"
'@lang VBA

Public LANGUAGE_ As String
Public COMMENT_LINE_ As String
Public COMMENT_MULTILINE_START_ As String
Public COMMENT_MULTILINE_END_ As String
Public STRING_MULTILINE_START_ As String
Public STRING_MULTILINE_END_ As String
Public RESERVED_ As Variant
Public OPERATORS_ As Variant
Public TYPES_ As Variant
Public BUILTINS_ As Variant
Public LITERALS_ As Variant

Public Sub HighlightSelection(ByVal func As String)
    Application.ScreenUpdating = False
        
    Dim Highlighter As clsHighlighter
    Set Highlighter = New clsHighlighter
    
    Application.Run func
    
    Highlighter.Language = LANGUAGE_
    Highlighter.CommentLine = COMMENT_LINE_
    Highlighter.CommentMultilineStart = COMMENT_MULTILINE_START_
    Highlighter.CommentMultilineEnd = COMMENT_MULTILINE_END_
    Highlighter.StringMultilineStart = STRING_MULTILINE_START_
    Highlighter.StringMultilineEnd = STRING_MULTILINE_END_
    Highlighter.Operators = OPERATORS_
    Highlighter.Reserved = RESERVED_
    Highlighter.Types = TYPES_
    Highlighter.Builtins = BUILTINS_
    Highlighter.Literals = LITERALS_

    '' Default format values that are initialized in class call but can be customized:
    ' Highlighter.Format.FontName = "Consolas"
    ' Highlighter.Format.FontSize = 9
    ' Highlighter.Format.SpaceBefore = 10
    ' Highlighter.Format.SpaceAfter = 10
    ' Highlighter.Format.BorderStyle = wdLineStyleSingle

    ' Clean range from previous formatting and apply new formatting
    Highlighter.PrepareRange
    
    ' Highlight range
    Highlighter.Highlight
    
    Selection.Collapse Direction:=wdCollapseEnd
    
End Sub

