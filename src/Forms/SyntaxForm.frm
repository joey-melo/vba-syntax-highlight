VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SyntaxForm 
   Caption         =   "Selector"
   ClientHeight    =   1449
   ClientLeft      =   100
   ClientTop       =   400
   ClientWidth     =   3800
   OleObjectBlob   =   "SyntaxForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SyntaxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()

    With SyntaxForm
        .Hide
        
        For i = 0 To UBound(LANGUAGES_)
            If .cbLanguageSelector.value = LANGUAGES_(i)(0) Then
                HighlightSelection LANGUAGES_(i)(1)
                Exit For
            End If
        Next i

        .cbLanguageSelector.Clear
        
    End With
End Sub

