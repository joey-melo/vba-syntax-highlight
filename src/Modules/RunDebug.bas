Attribute VB_Name = "RunDebug"
'@lang VBA

Sub RunDebug()
    ' This Sub is here for testing purposes so there is no reliance on the ribbon button
    ' Replace the func variable with the Sub you want to call.
    
    Dim func As String
    
    func = "HighlightCsharp"
    
    HighlightSelection func

End Sub

