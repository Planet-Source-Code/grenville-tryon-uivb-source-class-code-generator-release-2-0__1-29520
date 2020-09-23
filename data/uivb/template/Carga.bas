Attribute VB_Name = "Carga"

Public Arrini As Variant
Public cn As adodb.Connection

Public Sistema As Sistema
Public cliente As cliente


Public Function tsleeini(ByVal Archivo As String, Optional Completo As Boolean) As Variant
Dim Rpta() As Variant, texto As Variant, Actual As Double
Actual = 0
Open Archivo For Input As #1
Do While Not EOF(1)
    Line Input #1, texto
    If Mid(texto, 1, 1) <> ";" Then
        ReDim Preserve Rpta(Actual)
        Rpta(Actual) = texto
        Actual = Actual + 1
    Else
        If Completo Then
             ReDim Preserve Rpta(Actual)
             Rpta(Actual) = texto
            Actual = Actual + 1
        End If
    End If
Loop
tsleeini = Rpta
Close #1
End Function

'OBTIENE UNA VARIABLE DEL ARRAY INI
Public Function tsgetini(ByVal Arr As Variant, ByVal Variable As String, ByVal DEFAULT As String) As String
Dim Contador As Double, Indice As Double, Rpta As String, Actual As String, Buffer As String
Rpta = DEFAULT
For Contador = 0 To UBound(Arr)
    Actual = RTrim(UCase(Arr(Contador)))
    Buffer = RTrim(Arr(Contador))
    Variable = RTrim(UCase(Variable))
    Indice = InStr(1, Actual, Variable)
    If Indice = 1 And Mid(Buffer, Len(Variable) + 1, 1) = "=" Then
        Rpta = Mid(Buffer, Len(Variable) + 2)
        Exit For
    End If
Next
tsgetini = Rpta
End Function

'CONVIERTE UNA CADENA EN ARRAY
Public Function tsstrarr(ByVal Cadena As String, ByVal delimitador As String)
Dim Rpta() As Variant, Contador As String, Numero As Double, Buffer As String
Numero = 0
If IsEmpty(delimitador) Then
 delimitador = "|"
End If
Cadena = Cadena + delimitador
While InStr(1, Cadena, delimitador) <> 0
    Buffer = Mid(Cadena, 1, InStr(1, Cadena, delimitador) - 1)
    ReDim Preserve Rpta(Numero)
    Rpta(Numero) = Buffer
    Numero = Numero + 1
    Cadena = Mid(Cadena, InStr(1, Cadena, delimitador) + 1)
Wend
tsstrarr = Rpta
End Function

'GRABA EL ARCHIVO INI
Public Function tsgraini(Archivo As String, Variable As String, Valor As String, Optional Posicion As String)
Dim Arr As Variant, Indice As Double, Cadena As String, Rpta As Variant
If Len(Trim(Dir(Archivo))) > 0 Then
  Arr = tsleeini(Archivo, True)
  Variable = UCase(RTrim(Variable))
  For Contador = 0 To UBound(Arr)
      Cadena = UCase(RTrim(Arr(Contador)))
      If InStr(1, Cadena, Variable) = 1 Then
        Indice = Contador
      End If
  Next
  Open Archivo For Output As #1
  If Indice = 0 Then
     For Contador = 0 To UBound(Arr)
         If Len(Trim(Posicion)) > 0 Then
            If Arr(Contador) = Posicion Then
                 Print #1, Variable + "=" + Valor
            End If
         End If
         Print #1, Arr(Contador)
     Next
     If Len(Trim(Posicion)) = 0 Then
        Print #1, Variable + "=" + Valor
     End If
  Else
     For Contador = 0 To Indice - 1
         Print #1, Arr(Contador)
     Next
     If Len(Trim(Valor)) > 0 Then
        Print #1, LCase(Variable) + "=" + Valor
     End If
     For Contador = Indice + 1 To UBound(Arr)
         Print #1, Arr(Contador)
     Next
  End If
  Close #1
Else
  Open Archivo For Output As #1
   Print #1, Variable + "=" + Valor
  Close #1
End If
tsgraini = tsleeini(Archivo)
End Function


