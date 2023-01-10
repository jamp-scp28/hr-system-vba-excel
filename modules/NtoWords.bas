Attribute VB_Name = "NtoWords"
Function NumLetras(Valor As Currency, Optional MonedaSingular As String = "", Optional MonedaPlural As String = "") As String
Dim Cantidad As Currency, Centavos As Currency, Digito As Byte, PrimerDigito As Byte, SegundoDigito As Byte, TercerDigito As Byte, Bloque As String, NumeroBloques As Byte, BloqueCero
Dim Unidades As Variant, Decenas As Variant, Centenas As Variant, i As Variant 'Si esta como Option Explicit
Dim ValorEntero As Long
Dim ValorOriginal As Double
    Valor = Round(Valor, 2)
    Cantidad = Int(Valor)
    ValorEntero = Cantidad
    Centavos = (Valor - Cantidad) * 100
    Unidades = Array("UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
    Decenas = Array("DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    Centenas = Array("CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    NumeroBloques = 1
    
    Do
        PrimerDigito = 0
        SegundoDigito = 0
        TercerDigito = 0
        Bloque = ""
        BloqueCero = 0
        For i = 1 To 3
            Digito = Cantidad Mod 10
            If Digito <> 0 Then
                Select Case i
                Case 1
                    Bloque = " " & Unidades(Digito - 1)
                    PrimerDigito = Digito
                Case 2
                    If Digito <= 2 Then
                        Bloque = " " & Unidades((Digito * 10) + PrimerDigito - 1)
                    Else
                        Bloque = " " & Decenas(Digito - 1) & IIf(PrimerDigito <> 0, " Y", Null) & Bloque
                    End If
                    SegundoDigito = Digito
                Case 3
                    Bloque = " " & IIf(Digito = 1 And PrimerDigito = 0 And SegundoDigito = 0, "CIEN", Centenas(Digito - 1)) & Bloque
                    TercerDigito = Digito
                End Select
            Else
                BloqueCero = BloqueCero + 1
            End If
            Cantidad = Int(Cantidad / 10)
            If Cantidad = 0 Then
                Exit For
            End If
        Next i
        Select Case NumeroBloques
            Case 1
                NumLetras = Bloque
            Case 2
                NumLetras = Bloque & IIf(BloqueCero = 3, Null, " MIL") & NumLetras
            Case 3
                NumLetras = Bloque & IIf(PrimerDigito = 1 And SegundoDigito = 0 And TercerDigito = 0, " MILLON", " MILLONES") & NumLetras
        End Select
        NumeroBloques = NumeroBloques + 1
    Loop Until Cantidad = 0
    
    'Millardos
    If Valor >= 1000000000 Then
        Dim millardos As Currency
        Dim millarodsInt As Integer
        Dim letras_Millardos As String
        millarodsInt = Int(Valor / 1000000000)
        millardos = millarodsInt
        
        letras_Millardos = Replace(Trim(NumLetras(millardos)), "", IIf(millarodsInt = 1, "MILLARDO", "MILLARDOS"))
        NumLetras = letras_Millardos & NumLetras & "PESOS"
    End If
    
    NumLetras = Trim(NumLetras) & " PESOS"
End Function
Function NumTwords(Valor As Currency, Optional MonedaSingular As String = "", Optional MonedaPlural As String = "") As String
Dim Cantidad As Currency, Centavos As Currency, Digito As Byte, PrimerDigito As Byte, SegundoDigito As Byte, TercerDigito As Byte, Bloque As String, NumeroBloques As Byte, BloqueCero
Dim Unidades As Variant, Decenas As Variant, Centenas As Variant, i As Variant 'Si esta como Option Explicit
Dim ValorEntero As Long
Dim ValorOriginal As Double
    Valor = Round(Valor, 2)
    Cantidad = Int(Valor)
    ValorEntero = Cantidad
    Centavos = (Valor - Cantidad) * 100
    Unidades = Array("UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
    Decenas = Array("DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    Centenas = Array("CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    NumeroBloques = 1
    
    Do
        PrimerDigito = 0
        SegundoDigito = 0
        TercerDigito = 0
        Bloque = ""
        BloqueCero = 0
        For i = 1 To 3
            Digito = Cantidad Mod 10
            If Digito <> 0 Then
                Select Case i
                Case 1
                    Bloque = " " & Unidades(Digito - 1)
                    PrimerDigito = Digito
                Case 2
                    If Digito <= 2 Then
                        Bloque = " " & Unidades((Digito * 10) + PrimerDigito - 1)
                    Else
                        Bloque = " " & Decenas(Digito - 1) & IIf(PrimerDigito <> 0, " Y", Null) & Bloque
                    End If
                    SegundoDigito = Digito
                Case 3
                    Bloque = " " & IIf(Digito = 1 And PrimerDigito = 0 And SegundoDigito = 0, "CIEN", Centenas(Digito - 1)) & Bloque
                    TercerDigito = Digito
                End Select
            Else
                BloqueCero = BloqueCero + 1
            End If
            Cantidad = Int(Cantidad / 10)
            If Cantidad = 0 Then
                Exit For
            End If
        Next i
        Select Case NumeroBloques
            Case 1
                NumTwords = Bloque
            Case 2
                NumTwords = Bloque & IIf(BloqueCero = 3, Null, " MIL") & NumTwords
            Case 3
                NumTwords = Bloque & IIf(PrimerDigito = 1 And SegundoDigito = 0 And TercerDigito = 0, " MILLON", " MILLONES") & NumTwords
        End Select
        NumeroBloques = NumeroBloques + 1
    Loop Until Cantidad = 0
    
    'Millardos
    If Valor >= 1000000000 Then
        Dim millardos As Currency
        Dim millarodsInt As Integer
        Dim letras_Millardos As String
        millarodsInt = Int(Valor / 1000000000)
        millardos = millarodsInt
        
        letras_Millardos = Replace(Trim(NumLetras(millardos)), "", IIf(millarodsInt = 1, "MILLARDO", "MILLARDOS"))
        NumTwords = letras_Millardos & NumTwords
    End If
    
    NumTwords = StrConv(Trim(NumTwords), vbProperCase)
End Function
