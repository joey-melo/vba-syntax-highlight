VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private s_fontName As String
Private i_fontSize As Integer
Private i_spaceBefore As Integer
Private i_spaceAfter As Integer
Private o_borderStyle As WdLineStyle
Private c_color As WdColor
Private b_italics As Boolean

Property Get FontName() As String
    FontName = s_fontName
End Property
Property Let FontName(value As String)
    s_fontName = value
End Property
Property Get FontSize() As Integer
    FontSize = i_fontSize
End Property
Property Let FontSize(value As Integer)
    i_fontSize = value
End Property
Property Get SpaceBefore() As Integer
    SpaceBefore = i_spaceBefore
End Property
Property Let SpaceBefore(value As Integer)
    i_spaceBefore = value
End Property
Property Get SpaceAfter() As Integer
    SpaceAfter = i_spaceAfter
End Property
Property Let SpaceAfter(value As Integer)
    i_spaceAfter = value
End Property
Property Get BorderStyle() As String
    BorderStyle = o_borderStyle
End Property
Property Let BorderStyle(value As String)
    o_borderStyle = value
End Property
Property Get Color() As WdColor
    Color = c_color
End Property
Property Let Color(value As WdColor)
    c_color = value
End Property
Property Get Italics() As Boolean
    Italics = b_italics
End Property
Property Let Italics(value As Boolean)
    b_italics = value
End Property

Public Sub CleanRange(rng As Range)
    Selection.ClearFormatting

    ' Find/Replace does not work for this as it will bug the rest of the of Find/Replaces in the code.
    While True
loopStart:
        posTab = InStr(1, rng.text, Chr(9))
        posLineBreak = InStr(1, rng.text, Chr(11))
        posNonbreakingSpace = InStr(1, rng.text, Chr(160))
        posNonbreakingSpace2 = InStr(1, rng.text, Chr(176))
        posLeftDoubleQuote = InStr(1, rng.text, Chr(210))
        posRightDoubleQuote = InStr(1, rng.text, Chr(211))
        posLeftSingleQuote = InStr(1, rng.text, Chr(212))
        posRightSingleQuote = InStr(1, rng.text, Chr(213))
        
        If posTab > 0 Then
            rng.Characters(posTab) = "  " ' Replace tabs with double spaces (this is easier to read in a Word document)
            GoTo loopStart
        ElseIf posLineBreak > 0 Then
            rng.Characters(posLineBreak) = Chr(13) ' Replace line breaks with new line
            GoTo loopStart
        ElseIf posNonbreakingSpace > 0 Then
            rng.Characters(posNonbreakingSpace) = Chr(32) ' Replace non-breaking space with regular space
        ElseIf posNonbreakingSpace2 > 0 Then
            rng.Characters(posNonbreakingSpace2) = Chr(32) ' Replace non-breaking space with regular space
        ElseIf posLeftDoubleQuote > 0 Then
            rng.Characters(posLeftDoubleQuote) = Chr(34) ' Replace curly double quotes (left) with straight double quotes
            GoTo loopStart
        ElseIf posRightDoubleQuote > 0 Then
            rng.Characters(posRightDoubleQuote) = Chr(34) ' Replace curly double quotes (right) with straight double quotes
            GoTo loopStart
        ElseIf posLeftSingleQuote > 0 Then
            rng.Characters(posLeftSingleQuote) = Chr(39) ' Replace curly single quotes (left) with straight single quotes
            GoTo loopStart
        ElseIf posRightSingleQuote > 0 Then
            rng.Characters(posRightSingleQuote) = Chr(39) ' Replace curly single quotes (right) with straight single quotes
            GoTo loopStart
        End If
        Exit Sub
    Wend
End Sub
Public Sub ApplyBorders(rng As Range)
    Dim rngCopy As Range
    Set rngCopy = rng.Duplicate
    
    ' Borders in a single paragraph are messed up, we need to make sure there is a least two paragraphs in the selection
    rngCopy.Collapse Direction:=wdCollapseEnd
    rngCopy.InsertAfter text:=vbNewLine
    rng.End = rng.End + 1
    
    With rng.Borders
        .Item(wdBorderTop).LineStyle = o_borderStyle
        .Item(wdBorderBottom).LineStyle = o_borderStyle
        .Item(wdBorderLeft).LineStyle = o_borderStyle
        .Item(wdBorderRight).LineStyle = o_borderStyle
    End With
    
    ' Then delete the extra paragraph created
    rngCopy.Delete
    
End Sub
Public Sub ApplySpacing(rng As Range, spcBefore As Integer, spcAfter As Integer)
    rng.Paragraphs(1).Range.ParagraphFormat.SpaceBefore = spcBefore
    rng.Paragraphs(rng.Paragraphs.Count).Range.ParagraphFormat.SpaceAfter = spcAfter
End Sub
Public Sub SetFont(rng As Range, FontName As String, FontSize As Integer)
    With rng.Font
        .Name = FontName
        .size = FontSize
    End With
End Sub
Public Sub DisableProofing(rng As Range)
    rng.SpellingChecked = False
    rng.NoProofing = True
    rng.GrammarChecked = False
End Sub

