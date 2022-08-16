Attribute VB_Name = "TERZAGUI"
'Autor: Raimundo Leirton
'Data: 11/08/2022
'Descrisão: Calula o fator de forma da sapata para parte coesiva
Function gtFator_Sc(Forma_Sapata As String)
Attribute gtFator_Sc.VB_Description = "Determina o fator de forma Sc(relativo a coesão) para sapatas via a proposta de Terzagui"
Attribute gtFator_Sc.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error Resume Next
    gtFator_Sc = ""
    
    Select Case LCase(Forma_Sapata)
        Case "corrida":
            gtFator_Sc = 1
        Case "quadrada":
            gtFator_Sc = 1.3
        Case "circular":
            gtFator_Sc = 1.3
        Case "retangular":
            gtFator_Sc = 1.2
        End Select
End Function

'Autor: Raimundo Leirton
'Data: 11/08/2022
'Descrição: Calula o fator de forma da sapata para parte gravitacional
Function gtFator_Sg(Forma_Sapata As String)
Attribute gtFator_Sg.VB_Description = "Determina o fator de forma Sg(relativo a gravidade) para sapatas via a proposta de Terzagui"
Attribute gtFator_Sg.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error Resume Next
    gtFator_Sg = ""
    
    Select Case LCase(Forma_Sapata)
        Case "corrida":
            gtFator_Sg = 1
        Case "quadrada":
            gtFator_Sg = 0.8
        Case "circular":
            gtFator_Sg = 0.6
        Case "retangular":
            gtFator_Sg = 0.9
        End Select
End Function

'Autor: Raimundo Leirton
'Data: 11/08/2022
'Descrição: Calula o fator de forma da sapata para parte da carga da sapata
Function gtFator_Sq(Forma_Sapata As String)
Attribute gtFator_Sq.VB_Description = "Determina o fator de forma Sq(relativo a carga) para sapatas via a proposta de Terzagui"
Attribute gtFator_Sq.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error Resume Next
    gtFator_Sq = ""
    
    Select Case LCase(Forma_Sapata)
        Case "corrida":
            gtFator_Sq = 1
        Case "quadrada":
            gtFator_Sq = 1
        Case "circular":
            gtFator_Sq = 1
        Case "retangular":
            gtFator_Sq = 1
        End Select
End Function
