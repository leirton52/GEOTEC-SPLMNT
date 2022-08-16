Attribute VB_Name = "INFO_FUNCOES"
'Autor: Raimundo
'Data: 10/08/2022
'Descrição: Cria descrições para as funções criadas

Sub Gerar_Info_Funcoes()
    On Error Resume Next
    
    Dim Arg(0) As String 'Argunmentos da função
    
    With Application
        'Inserir comentários na função gtCOMPACIDAE
        Arg(0) = "Indice de prenetração do ensaio spt"
        .MacroOptions "gtCOMPACIDAE", _
            "Determina a compacidade de solos arenos via a proposta de Terzagui e Pack(1948)", _
            , , , , _
            "GEOTECNIA", _
            , , , Arg
        
        'Inserir comentários na função CONSISTÊNCIA
        Arg(0) = "Indice de prenetração do ensaio spt"
        .MacroOptions "gtCONSISTENCIA", _
            "Determina a consistência de solos argilosos via a proposta de Terzagui e Pack(1948)", _
            , , , , _
            "GEOTECNIA", _
            , , , Arg
    End With
End Sub

