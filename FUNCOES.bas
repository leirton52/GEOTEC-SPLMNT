Attribute VB_Name = "FUNCOES"
'Autor: Raimundo
'Data: 05/08/2022
'
'Descriss�o: Calculando a �rea de uma bitola com o di�metro

Public Function A_CIRC(Optional diametro As Variant)
    Pi = 3.14
    A_CIRC = ""
    
    'Valida se o di�metro foi inserido e � um n�mero
    If ((Not IsNumeric(diametro)) Or diametro = "") Then: Exit Function
    
    'calcula a �rea da sec��o transversal da bitola
    A_CIRC = Pi * diametro ^ 2 / 4
End Function
    
'Autor: Raimundo
'Data: 05/08/2022
'
'Descriss�o: Verifica a consist�ncia de uma argila de acordo com o Nspt

Public Function gtCONSISTENCIA(Optional N_spt As Variant)
Attribute gtCONSISTENCIA.VB_Description = "Determina a consist�ncia de solos argilosos via a proposta de Terzagui e Pack(1948)"
Attribute gtCONSISTENCIA.VB_ProcData.VB_Invoke_Func = " \n20"
    gtCONSISTENCIA = ""
    
    'Valida se o N_spt foi inserido e � um n�mero
    If ((Not IsNumeric(N_spt)) Or N_spt = "") Then Exit Function
    'Valida se o N_spt foi inserido e � um n�mero
    If (N_spt <= 0) Then Exit Function
    
    If N_spt < 2 Then
        gtCONSISTENCIA = "Muito mole"
    ElseIf N_spt <= 4 Then
        gtCONSISTENCIA = "Mole"
    ElseIf N_spt <= 8 Then
        gtCONSISTENCIA = "M�dia"
    ElseIf N_spt <= 15 Then
        gtCONSISTENCIA = "Dura"
    ElseIf N_spt <= 30 Then
        gtCONSISTENCIA = "Muito Dura"
    Else
        gtCONSISTENCIA = "Rija"
    End If
End Function

'Autor: Raimundo
'Data: 05/08/2022
'
'Descriss�o: Verifica a consist�ncia de uma argila de acordo com o Nspt

Public Function gtCOMPACIDAE(Optional N_spt As Variant)
Attribute gtCOMPACIDAE.VB_Description = "Determina a compacidade de solos arenos via a proposta de Terzagui e Pack(1948)"
Attribute gtCOMPACIDAE.VB_ProcData.VB_Invoke_Func = " \n20"
    gtCOMPACIDAE = ""
    
    'Valida se o N_spt foi inserido e � um n�mero
    If ((Not IsNumeric(N_spt)) Or N_spt = "") Then Exit Function
    'Valida se o N_spt foi inserido e � um n�mero
    If (N_spt <= 0) Then Exit Function
    
    If N_spt < 4 Then
        gtCOMPACIDAE = "Muito solta"
    ElseIf N_spt <= 10 Then
        gtCOMPACIDAE = "Solta"
    ElseIf N_spt <= 30 Then
        gtCOMPACIDAE = "Medianamente densa"
    ElseIf N_spt <= 50 Then
        gtCOMPACIDAE = "Densa"
    Else
        gtCOMPACIDAE = "Muito densa"
    End If
End Function

