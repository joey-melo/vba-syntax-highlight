VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private o_literals As clsFormat
Private o_reserved As clsFormat
Private o_builtins As clsFormat
Private o_types As clsFormat
Private o_functionCalls As clsFormat
Private o_strings As clsFormat
Private o_comments As clsFormat
Private o_functionParameters As clsFormat
Private o_operators As clsFormat
Private o_numbers As clsFormat
Private o_cmdArguments As clsFormat
Private o_variables As clsFormat
Private o_tags As clsFormat

Property Get KeywordsLiterals() As clsFormat
   Set KeywordsLiterals = o_literals
End Property
Property Set KeywordsLiterals(value As clsFormat)
    o_literals = value
End Property
Property Get KeywordsReserved() As clsFormat
   Set KeywordsReserved = o_reserved
End Property
Property Set KeywordsReserved(value As clsFormat)
    o_reserved = value
End Property
Property Get KeywordsBuiltins() As clsFormat
   Set KeywordsBuiltins = o_builtins
End Property
Property Set KeywordsBuiltins(value As clsFormat)
    o_builtins = value
End Property
Property Get KeywordsTypes() As clsFormat
   Set KeywordsTypes = o_types
End Property
Property Set KeywordsTypes(value As clsFormat)
    o_types = value
End Property
Property Get FunctionCalls() As clsFormat
   Set FunctionCalls = o_functionCalls
End Property
Property Set FunctionCalls(value As clsFormat)
    o_functionCalls = value
End Property
Property Get Strings() As clsFormat
   Set Strings = o_strings
End Property
Property Set Strings(value As clsFormat)
    o_strings = value
End Property
Property Get Comments() As clsFormat
   Set Comments = o_comments
End Property
Property Set Comments(value As clsFormat)
    o_comments = value
End Property
Property Get FunctionParameters() As clsFormat
   Set FunctionParameters = o_functionParameters
End Property
Property Set FunctionParameters(value As clsFormat)
    o_functionParameters = value
End Property
Property Get Operators() As clsFormat
   Set Operators = o_operators
End Property
Property Set Operators(value As clsFormat)
    o_operators = value
End Property
Property Get Numbers() As clsFormat
   Set Numbers = o_numbers
End Property
Property Set Numbers(value As clsFormat)
    o_numbers = value
End Property
Property Get CmdArguments() As clsFormat
   Set CmdArguments = o_cmdArguments
End Property
Property Set CmdArguments(value As clsFormat)
    o_cmdArguments = value
End Property
Property Get Variables() As clsFormat
   Set Variables = o_variables
End Property
Property Set Variables(value As clsFormat)
    o_variables = value
End Property
Property Get Tags() As clsFormat
   Set Tags = o_tags
End Property
Property Set Tags(value As clsFormat)
    o_tags = value
End Property

Private Sub Class_Initialize()
    Set o_literals = New clsFormat
    Set o_reserved = New clsFormat
    Set o_builtins = New clsFormat
    Set o_types = New clsFormat
    Set o_functionCalls = New clsFormat
    Set o_strings = New clsFormat
    Set o_comments = New clsFormat
    Set o_functionParameters = New clsFormat
    Set o_operators = New clsFormat
    Set o_numbers = New clsFormat
    Set o_cmdArguments = New clsFormat
    Set o_variables = New clsFormat
    Set o_tags = New clsFormat
    
    With o_literals
        .Color = wdColorRed
        .Italics = True
    End With
    
    With o_reserved
    .Color = wdColorBlue
    .Italics = False
    End With
     
    With o_builtins
        .Color = wdColorBlue
        .Italics = True
    End With
     
    With o_types
        .Color = wdColorBlue
        .Italics = True
    End With
     
    With o_functionCalls
        .Color = wdColorTeal
        .Italics = False
    End With
     
    With o_strings
        .Color = wdColorGreen
        .Italics = False
    End With
     
    With o_comments
        .Color = wdColorGray50
        .Italics = False
    End With
     
    With o_functionParameters
        .Color = wdColorBrown
        .Italics = False
    End With
     
    With o_operators
        .Color = wdColorOrange
        .Italics = False
    End With
     
    With o_numbers
        .Color = wdColorDarkYellow
        .Italics = False
    End With
     
    With o_cmdArguments
        .Color = wdColorViolet
        .Italics = False
    End With
     
    With o_variables
        .Color = wdColorOrange
        .Italics = False
    End With
    
    With o_tags
        .Color = wdColorRed
        .Italics = False
    End With

End Sub

