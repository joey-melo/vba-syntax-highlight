Attribute VB_Name = "FormHandler"
'@lang VBA

Public LANGUAGES_ As Variant
''

Public Sub SyntaxFormHandler()
    Dim selector As SyntaxForm
    Set selector = New SyntaxForm
    
    ' Map languages to their function calls
    ' Reverse alphabetical order here = alphabetical order in form display
    LANGUAGES_ = Array( _
        Array("Xml", "HighlightXml"), _
        Array("Shell", "HighlightShell"), _
        Array("Python", "HighlightPython"), _
        Array("Json", "HighlightJson"), _
        Array("Java", "HighlightJava"), _
        Array("Html", "HighlightHtml"), _
        Array("CSharp", "HighlightCsharp"), _
        Array("C++", "HighlightCpp"), _
        Array("C", "HighlightC") _
    )

    With SyntaxForm
        ' Load selection box
        With .cbLanguageSelector
            For i = 0 To UBound(LANGUAGES_)
                .AddItem LANGUAGES_(i)(0), 0
            Next i
        End With
        
        .Show
    End With
End Sub



