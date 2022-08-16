Attribute VB_Name = "uteis"
Sub Converter_Maiuscaulas()
    'converte as celulas selecionadas em letras maiúsculas
    
    Dim Cel As Range
    
    For Each Cel In Selection
        Cel.Value = UCase(Cel.Value)
    Next
End Sub
