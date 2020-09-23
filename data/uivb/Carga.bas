Attribute VB_Name = "Carga"

Public cn As ADODB.Connection

Public Sub CreaRuta(NombreArchivo As String)
Dim Contador As Integer, Buffer As String
On Error Resume Next
NombreArchivo = Trim(NombreArchivo)
If Mid(NombreArchivo, Len(NombreArchivo), 1) <> "\" Then
    NombreArchivo = NombreArchivo + "\"
End If
Buffer = ""
For Contador = 1 To Len(NombreArchivo)
    If Mid(NombreArchivo, Contador, 1) = "\" Then
        MkDir Buffer
    End If
    Buffer = Buffer + Mid(NombreArchivo, Contador, 1)
Next
On Error GoTo 0
End Sub

Public Function StrTran(ByVal Cadena As String, ByVal Inicial As String, ByVal Final As String) As String
Dim Contador As Double, Buffer As String
Contador = 0
Do While True
    Contador = Contador + 1
    If Contador > Len(Cadena) Then
        Exit Do
    End If
    If Mid(Cadena, Contador, Len(Inicial)) = Inicial Then
        Buffer = Mid(Cadena, 1, Contador - 1) + Final + Mid(Cadena, Contador + Len(Inicial))
        Cadena = Buffer
    End If
Loop
StrTran = Cadena
End Function

Public Function FCase(Texto) As String
Texto = LCase(Texto)
FCase = UCase(Mid(Texto, 1, 1)) + Mid(Texto, 2)
End Function

