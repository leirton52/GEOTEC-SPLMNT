Attribute VB_Name = "INFO_FUNCOES"
'Autor: Raimundo
'Data: 10/08/2022
'Descri��o: Cria descri��es para as fun��es criadas

Sub Gerar_Info_Funcoes()
    On Error Resume Next
    
    Dim Arg(0) As String 'Argunmentos da fun��o
    
    With Application
        'Inserir coment�rios na fun��o gtCOMPACIDAE
        Arg(0) = "Indice de prenetra��o do ensaio spt"
        .MacroOptions "gtCOMPACIDAE", _
            "Determina a compacidade de solos arenos via a proposta de Terzagui e Pack(1948)", _
            , , , , _
            "GEOTECNIA", _
            , , , Arg
        
        'Inserir coment�rios na fun��o CONSIST�NCIA
        Arg(0) = "Indice de prenetra��o do ensaio spt"
        .MacroOptions "gtCONSISTENCIA", _
            "Determina a consist�ncia de solos argilosos via a proposta de Terzagui e Pack(1948)", _
            , , , , _
            "GEOTECNIA", _
            , , , Arg
    End With
End Sub

