Attribute VB_Name = "RibbonRouter"
'@lang VBA

Sub ApplyPythonSyntaxHighlight(control As IRibbonControl)
    
    HighlightSelection "HighlightPython"

End Sub
Sub ApplyShellSyntaxHighlight(control As IRibbonControl)

    HighlightSelection "HighlightShell"

End Sub
Sub ApplyHtmlSyntaxHighlight(control As IRibbonControl)
   
   HighlightSelection "HighlightHtml"
  
End Sub
Sub ApplyCustomSyntaxHighlight(control As IRibbonControl)

    SyntaxFormHandler
    
End Sub

