Attribute VB_Name = "Plantillas"

'OK
Public Sub CreaCarga()
Dim Contador As Integer, Nombre As String
Dim Cadena As String
Cadena = ""
Cadena = Cadena + "Attribute VB_Name = ""Carga"" " + vbCrLf
Cadena = Cadena + " " + vbCrLf
Cadena = Cadena + "Public Arrini As Variant" + vbCrLf
Cadena = Cadena + "Public cn As adodb.Connection" + vbCrLf
Cadena = Cadena + "" + vbCrLf
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count ' - 1
    If frmPrincipal.tre(0).Nodes(Contador).Checked And frmPrincipal.tre(0).Nodes(Contador).Children > 0 Then
        Cadena = Cadena + "Public " + FCase(frmPrincipal.tre(0).Nodes(Contador).Text) + " As " + FCase(frmPrincipal.tre(0).Nodes(Contador).Text) + vbCrLf
    End If
Next
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Public Function tsleeini(ByVal Archivo As String, Optional Completo As Boolean) As Variant" + vbCrLf
Cadena = Cadena + "Dim Rpta() As Variant, texto As Variant, Actual As Double" + vbCrLf
Cadena = Cadena + "Actual = 0" + vbCrLf
Cadena = Cadena + "Open Archivo For Input As #1" + vbCrLf
Cadena = Cadena + "Do While Not EOF(1)" + vbCrLf
Cadena = Cadena + "    Line Input #1, texto" + vbCrLf
Cadena = Cadena + "    If Mid(texto, 1, 1) <> "";"" Then" + vbCrLf
Cadena = Cadena + "        ReDim Preserve Rpta(Actual)" + vbCrLf
Cadena = Cadena + "        Rpta(Actual) = texto" + vbCrLf
Cadena = Cadena + "        Actual = Actual + 1" + vbCrLf
Cadena = Cadena + "    Else" + vbCrLf
Cadena = Cadena + "        If Completo Then" + vbCrLf
Cadena = Cadena + "             ReDim Preserve Rpta(Actual)" + vbCrLf
Cadena = Cadena + "             Rpta(Actual) = texto" + vbCrLf
Cadena = Cadena + "            Actual = Actual + 1" + vbCrLf
Cadena = Cadena + "        End If" + vbCrLf
Cadena = Cadena + "    End If" + vbCrLf
Cadena = Cadena + "Loop" + vbCrLf
Cadena = Cadena + "tsleeini = Rpta" + vbCrLf
Cadena = Cadena + "Close #1" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'OBTIENE UNA VARIABLE DEL ARRAY INI" + vbCrLf
Cadena = Cadena + "Public Function tsgetini(ByVal Arr As Variant, ByVal Variable As String, ByVal DEFAULT As String) As String" + vbCrLf
Cadena = Cadena + "Dim Contador As Double, Indice As Double, Rpta As String, Actual As String, Buffer As String" + vbCrLf
Cadena = Cadena + "Rpta = DEFAULT" + vbCrLf
Cadena = Cadena + "For Contador = 0 To UBound(Arr)" + vbCrLf
Cadena = Cadena + "    Actual = RTrim(UCase(Arr(Contador)))" + vbCrLf
Cadena = Cadena + "    Buffer = RTrim(Arr(Contador))" + vbCrLf
Cadena = Cadena + "    Variable = RTrim(UCase(Variable))" + vbCrLf
Cadena = Cadena + "    Indice = InStr(1, Actual, Variable)" + vbCrLf
Cadena = Cadena + "    If Indice = 1 And Mid(Buffer, Len(Variable) + 1, 1) = ""="" Then" + vbCrLf
Cadena = Cadena + "        Rpta = Mid(Buffer, Len(Variable) + 2)" + vbCrLf
Cadena = Cadena + "        Exit For" + vbCrLf
Cadena = Cadena + "    End If" + vbCrLf
Cadena = Cadena + "Next" + vbCrLf
Cadena = Cadena + "tsgetini = Rpta" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'CONVIERTE UNA CADENA EN ARRAY" + vbCrLf
Cadena = Cadena + "Public Function tsstrarr(ByVal Cadena As String, ByVal delimitador As String)" + vbCrLf
Cadena = Cadena + "Dim Rpta() As Variant, Contador As String, Numero As Double, Buffer As String" + vbCrLf
Cadena = Cadena + "Numero = 0" + vbCrLf
Cadena = Cadena + "If IsEmpty(delimitador) Then" + vbCrLf
Cadena = Cadena + " delimitador = ""|""" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "Cadena = Cadena + delimitador" + vbCrLf
Cadena = Cadena + "While InStr(1, Cadena, delimitador) <> 0" + vbCrLf
Cadena = Cadena + "    Buffer = Mid(Cadena, 1, InStr(1, Cadena, delimitador) - 1)" + vbCrLf
Cadena = Cadena + "    ReDim Preserve Rpta(Numero)" + vbCrLf
Cadena = Cadena + "    Rpta(Numero) = Buffer" + vbCrLf
Cadena = Cadena + "    Numero = Numero + 1" + vbCrLf
Cadena = Cadena + "    Cadena = Mid(Cadena, InStr(1, Cadena, delimitador) + 1)" + vbCrLf
Cadena = Cadena + "Wend" + vbCrLf
Cadena = Cadena + "tsstrarr = Rpta" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'GRABA EL ARCHIVO INI" + vbCrLf
Cadena = Cadena + "Public Function tsgraini(Archivo As String, Variable As String, Valor As String, Optional Posicion As String)" + vbCrLf
Cadena = Cadena + "Dim Arr As Variant, Indice As Double, Cadena As String, Rpta As Variant" + vbCrLf
Cadena = Cadena + "If Len(Trim(Dir(Archivo))) > 0 Then" + vbCrLf
Cadena = Cadena + "  Arr = tsleeini(Archivo, True)" + vbCrLf
Cadena = Cadena + "  Variable = UCase(RTrim(Variable))" + vbCrLf
Cadena = Cadena + "  For Contador = 0 To UBound(Arr)" + vbCrLf
Cadena = Cadena + "      Cadena = UCase(RTrim(Arr(Contador)))" + vbCrLf
Cadena = Cadena + "      If InStr(1, Cadena, Variable) = 1 Then" + vbCrLf
Cadena = Cadena + "        Indice = Contador" + vbCrLf
Cadena = Cadena + "      End If" + vbCrLf
Cadena = Cadena + "  Next" + vbCrLf
Cadena = Cadena + "  Open Archivo For Output As #1" + vbCrLf
Cadena = Cadena + "  If Indice = 0 Then" + vbCrLf
Cadena = Cadena + "     For Contador = 0 To UBound(Arr)" + vbCrLf
Cadena = Cadena + "         If Len(Trim(Posicion)) > 0 Then" + vbCrLf
Cadena = Cadena + "            If Arr(Contador) = Posicion Then" + vbCrLf
Cadena = Cadena + "                 Print #1, Variable + ""="" + Valor" + vbCrLf
Cadena = Cadena + "            End If" + vbCrLf
Cadena = Cadena + "         End If" + vbCrLf
Cadena = Cadena + "         Print #1, Arr(Contador)" + vbCrLf
Cadena = Cadena + "     Next" + vbCrLf
Cadena = Cadena + "     If Len(Trim(Posicion)) = 0 Then" + vbCrLf
Cadena = Cadena + "        Print #1, Variable + ""="" + Valor" + vbCrLf
Cadena = Cadena + "     End If" + vbCrLf
Cadena = Cadena + "  Else" + vbCrLf
Cadena = Cadena + "     For Contador = 0 To Indice - 1" + vbCrLf
Cadena = Cadena + "         Print #1, Arr(Contador)" + vbCrLf
Cadena = Cadena + "     Next" + vbCrLf
Cadena = Cadena + "     If Len(Trim(Valor)) > 0 Then" + vbCrLf
Cadena = Cadena + "        Print #1, LCase(Variable) + ""="" + Valor" + vbCrLf
Cadena = Cadena + "     End If" + vbCrLf
Cadena = Cadena + "     For Contador = Indice + 1 To UBound(Arr)" + vbCrLf
Cadena = Cadena + "         Print #1, Arr(Contador)" + vbCrLf
Cadena = Cadena + "     Next" + vbCrLf
Cadena = Cadena + "  End If" + vbCrLf
Cadena = Cadena + "  Close #1" + vbCrLf
Cadena = Cadena + "Else" + vbCrLf
Cadena = Cadena + "  Open Archivo For Output As #1" + vbCrLf
Cadena = Cadena + "   Print #1, Variable + ""="" + Valor" + vbCrLf
Cadena = Cadena + "  Close #1" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "tsgraini = tsleeini(Archivo)" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Open frmPrincipal.txt(1).Text + "carga.bas" For Output As #1
Print #1, Cadena
Close #1
End Sub

'OK
Public Function CreaMDI()
Dim Cadena As String, Contador As Integer, Nombre As String, Cuantos As Integer
Cadena = ""
Cadena = Cadena + "VERSION 5.00" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Begin VB.MDIForm MDI " + vbCrLf
Cadena = Cadena + "   BackColor       =   &H8000000C&" + vbCrLf
Cadena = Cadena + "   Caption         =   """ + frmPrincipal.txt(3).Text + """" + vbCrLf
Cadena = Cadena + "   ClientHeight    =   3195" + vbCrLf
Cadena = Cadena + "   ClientLeft      =   60" + vbCrLf
Cadena = Cadena + "   ClientTop       =   630" + vbCrLf
Cadena = Cadena + "   ClientWidth     =   4680" + vbCrLf
Cadena = Cadena + "   LinkTopic       =   ""MDIForm1""" + vbCrLf
Cadena = Cadena + "   StartUpPosition =   2  'CenterScreen" + vbCrLf
Cadena = Cadena + "   WindowState     =   2  'Maximized" + vbCrLf
Cadena = Cadena + "   Begin VB.Menu MnuArchivos " + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Archivos""" + vbCrLf
Cuantos = 0
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count ' - 1
    If Marcado(frmPrincipal.tre(0).Nodes(Contador).Text) And frmPrincipal.tre(0).Nodes(Contador).Children > 0 Then
        Cadena = Cadena + "      Begin VB.Menu MnuArchivo " + vbCrLf
        Cadena = Cadena + "         Caption         =   """ + "Mantenimiento de &" + FCase(frmPrincipal.tre(0).Nodes(Contador).Text) + """" + vbCrLf
        Cadena = Cadena + "         Index           =   " + CStr(Cuantos) + vbCrLf
        Cadena = Cadena + "      End" + vbCrLf
        Cuantos = Cuantos + 1
    End If
Next
Cadena = Cadena + "      Begin VB.Menu MnuArchivo " + vbCrLf
Cadena = Cadena + "         Caption         =   ""-""" + vbCrLf
Cadena = Cadena + "         Index           =   " + CStr(Cuantos) + vbCrLf
Cadena = Cadena + "      End" + vbCrLf
Cadena = Cadena + "      Begin VB.Menu MnuArchivo " + vbCrLf
Cadena = Cadena + "         Caption         =   ""&Salir""" + vbCrLf
Cadena = Cadena + "         Index           =   " + CStr(Cuantos + 1) + vbCrLf
Cadena = Cadena + "      End" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "End" + vbCrLf
Cadena = Cadena + "Attribute VB_Name = ""MDI""" + vbCrLf
Cadena = Cadena + "Attribute VB_GlobalNameSpace = False" + vbCrLf
Cadena = Cadena + "Attribute VB_Creatable = False" + vbCrLf
Cadena = Cadena + "Attribute VB_PredeclaredId = True" + vbCrLf
Cadena = Cadena + "Attribute VB_Exposed = False" + vbCrLf
Cadena = Cadena + "Private Sub MDIForm_Load()" + vbCrLf
Cadena = Cadena + "If Dir(App.Path + ""\"" + App.EXEName + "".ini"") <> """" Then" + vbCrLf
Cadena = Cadena + "    Arrini = tsleeini(App.Path + ""\"" + App.EXEName + "".ini"")" + vbCrLf
Cadena = Cadena + "Else" + vbCrLf
Cadena = Cadena + "    MsgBox ""ERROR: No se encuentra el archivo de configuracion """""" + App.Path + ""\"" + App.EXEName + "".ini"""". Se cancela el Proceso"", vbCritical + vbOKOnly, ""ERROR""" + vbCrLf
Cadena = Cadena + "    End" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "CargaClases" + vbCrLf
Cadena = Cadena + "If Not AbreConexion Then" + vbCrLf
Cadena = Cadena + "    MsgBox ""ERROR: No se puede crear conexion "" + cn.ConnectionString + "". Se cancela el Proceso"", vbCritical + vbOKOnly, ""ERROR""" + vbCrLf
Cadena = Cadena + "    End" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)" + vbCrLf
Cadena = Cadena + "Cancel = 1" + vbCrLf
Cadena = Cadena + "If MsgBox(""¿Desea salir del Software?"", vbYesNo + vbQuestion, ""ATENCION"") = vbYes Then" + vbCrLf
Cadena = Cadena + "    Cancel = 0" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub MnuArchivo_Click(Index As Integer)" + vbCrLf
Cadena = Cadena + "Select Case Index" + vbCrLf
Cuantos = 0
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count ' - 1
    If frmPrincipal.tre(0).Nodes(Contador).Checked And frmPrincipal.tre(0).Nodes(Contador).Children > 0 Then
        Nombre = frmPrincipal.tre(0).Nodes(Contador).Text
        If Marcado(Nombre) Then
            Cadena = Cadena + "Case " + CStr(Cuantos) + vbCrLf
            Cadena = Cadena + "    frm" + Nombre + ".Show" + vbCrLf
            Cuantos = Cuantos + 1
        End If
    End If
Next
Cadena = Cadena + "Case " + CStr(Cuantos + 1) + vbCrLf
Cadena = Cadena + "    Unload Me" + vbCrLf
Cadena = Cadena + "End Select" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub CargaClases()" + vbCrLf
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count ' - 1
    If frmPrincipal.tre(0).Nodes(Contador).Checked And frmPrincipal.tre(0).Nodes(Contador).Children > 0 Then
        Cadena = Cadena + "Set " + frmPrincipal.tre(0).Nodes(Contador).Text + " = New " + frmPrincipal.tre(0).Nodes(Contador).Text + vbCrLf
    End If
Next
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Function AbreConexion() As Boolean" + vbCrLf
Cadena = Cadena + "AbreConexion = False" + vbCrLf
Cadena = Cadena + "Set cn = New adodb.Connection" + vbCrLf
Cadena = Cadena + "cn.ConnectionString = tsgetini(Arrini, ""conexion"", """")" + vbCrLf
Cadena = Cadena + "cn.Open" + vbCrLf
Cadena = Cadena + "AbreConexion = True" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Open frmPrincipal.txt(1).Text + "mdi.frm" For Output As #1
Print #1, Cadena
Close #1
End Function

'OK
Public Sub CreaProyecto()
Dim Cadena As String, Contador As Integer, Nombre As String, Cuantos As Integer
Cadena = ""
Cadena = Cadena + "Type=Exe" + vbCrLf
Cadena = Cadena + "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\SYSTEM\stdole2.tlb#OLE Automation" + vbCrLf
Cadena = Cadena + "Reference=*\G{00000200-0000-0010-8000-00AA006D2EA4}#2.0#0#C:\ARCHIVOS DE PROGRAMA\ARCHIVOS COMUNES\SYSTEM\ADO\msado20.tlb#Microsoft ActiveX Data Objects 2.0 Library" + vbCrLf
Cadena = Cadena + "Form=MDI.frm" + vbCrLf
Cadena = Cadena + "Module=Carga; Carga.bas" + vbCrLf
Cadena = Cadena + "Object={5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0; MSFLXGRD.OCX" + vbCrLf
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count ' - 1
    If frmPrincipal.tre(0).Nodes(Contador).Checked And frmPrincipal.tre(0).Nodes(Contador).Children > 0 Then
        Nombre = FCase(frmPrincipal.tre(0).Nodes(Contador).Text)
        If Marcado(Nombre) Then
            Cadena = Cadena + "Form=frm" + Nombre + ".frm" + vbCrLf
            Cadena = Cadena + "Form=frmEdita" + Nombre + ".frm" + vbCrLf
        End If
        Cadena = Cadena + "Class=" + Nombre + "; " + Nombre + ".cls" + vbCrLf
    End If
Next
Cadena = Cadena + "Startup=""MDI""" + vbCrLf
Cadena = Cadena + "HelpFile=""""" + vbCrLf
Cadena = Cadena + "Command32=""""" + vbCrLf
Cadena = Cadena + "Name=""" + frmPrincipal.txt(2).Text + """" + vbCrLf
Cadena = Cadena + "HelpContextID=""0""" + vbCrLf
Cadena = Cadena + "CompatibleMode=""0""" + vbCrLf
Cadena = Cadena + "MajorVer=1" + vbCrLf
Cadena = Cadena + "MinorVer=0" + vbCrLf
Cadena = Cadena + "RevisionVer=0" + vbCrLf
Cadena = Cadena + "AutoIncrementVer=0" + vbCrLf
Cadena = Cadena + "ServerSupportFiles=0" + vbCrLf
Cadena = Cadena + "VersionCompanyName=""AFP HORIZONTE""" + vbCrLf
Cadena = Cadena + "CompilationType=0" + vbCrLf
Cadena = Cadena + "OptimizationType=0" + vbCrLf
Cadena = Cadena + "FavorPentiumPro(tm)=0" + vbCrLf
Cadena = Cadena + "CodeViewDebugInfo=0" + vbCrLf
Cadena = Cadena + "NoAliasing=0" + vbCrLf
Cadena = Cadena + "BoundsCheck=0" + vbCrLf
Cadena = Cadena + "OverflowCheck=0" + vbCrLf
Cadena = Cadena + "FlPointCheck=0" + vbCrLf
Cadena = Cadena + "FDIVCheck=0" + vbCrLf
Cadena = Cadena + "UnroundedFP=0" + vbCrLf
Cadena = Cadena + "StartMode=0" + vbCrLf
Cadena = Cadena + "Unattended=0" + vbCrLf
Cadena = Cadena + "Retained=0" + vbCrLf
Cadena = Cadena + "ThreadPerObject=0" + vbCrLf
Cadena = Cadena + "MaxNumberOfThreads=1" + vbCrLf
Open frmPrincipal.txt(1).Text + frmPrincipal.txt(2).Text + ".vbp" For Output As #1
Print #1, Cadena
Close #1
End Sub

Public Function CreaClase()
Dim Contador As Integer, Nombre As String, NombreClase As String
Dim Crear As Boolean, Arr As Variant
Arr = Array()
Crear = False
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count ' - 1
    If frmPrincipal.tre(0).Nodes(Contador).Children > 0 Then
        If Crear Then
            frmPrincipal.lst(1).AddItem "   - Clase " + NombreClase: DoEvents
            frmPrincipal.lst(1).Selected(frmPrincipal.lst(1).ListCount - 1) = True: DoEvents
            Creacion NombreClase, Arr
            Crear = False
            Arr = Array()
        End If
    End If
    If frmPrincipal.tre(0).Nodes(Contador).Children > 0 And frmPrincipal.tre(0).Nodes(Contador).Checked Then
        NombreClase = frmPrincipal.tre(0).Nodes(Contador).Text
        Crear = True
    End If
    If frmPrincipal.tre(0).Nodes(Contador).Children = 0 And Crear Then
        ReDim Preserve Arr(IIf(UBound(Arr, 1) = -1, 1, UBound(Arr, 1) + 1))
        Arr(UBound(Arr, 1) - 1) = frmPrincipal.tre(0).Nodes(Contador).Text + IIf(frmPrincipal.tre(0).Nodes(Contador).Checked, "*", "")
    End If
Next
If Crear Then
    frmPrincipal.lst(1).AddItem "   - Clase " + NombreClase: DoEvents
    frmPrincipal.lst(1).Selected(frmPrincipal.lst(1).ListCount - 1) = True: DoEvents
    Creacion NombreClase, Arr
End If
End Function

Public Function CreaMantenimiento()
Dim Contador As Integer, Nombre As String, NombreClase As String
Dim Crear As Boolean, Arr As Variant
Arr = Array()
Crear = False
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count
    If frmPrincipal.tre(0).Nodes(Contador).Children > 0 Then
        If Crear Then
            If Marcado(NombreClase) Then
                frmPrincipal.lst(1).AddItem "   - Mantenimiento " + NombreClase: DoEvents
                frmPrincipal.lst(1).Selected(frmPrincipal.lst(1).ListCount - 1) = True: DoEvents
                CreacionM NombreClase, Arr
                frmPrincipal.lst(1).AddItem "   - Edicion " + NombreClase: DoEvents
                frmPrincipal.lst(1).Selected(frmPrincipal.lst(1).ListCount - 1) = True: DoEvents
                CreacionE NombreClase, Arr
                Crear = False
            End If
            Arr = Array()
        End If
    End If
    If frmPrincipal.tre(0).Nodes(Contador).Children > 0 And frmPrincipal.tre(0).Nodes(Contador).Checked Then
        NombreClase = frmPrincipal.tre(0).Nodes(Contador).Text
        Crear = True
    End If
    If frmPrincipal.tre(0).Nodes(Contador).Children = 0 And Crear Then
        ReDim Preserve Arr(IIf(UBound(Arr, 1) = -1, 1, UBound(Arr, 1) + 1))
        Arr(UBound(Arr, 1) - 1) = frmPrincipal.tre(0).Nodes(Contador).Text + IIf(frmPrincipal.tre(0).Nodes(Contador).Checked, "*", "")
    End If
Next
If Crear And Marcado(NombreClase) Then
    frmPrincipal.lst(1).AddItem "   - Mantenimiento " + NombreClase: DoEvents
    frmPrincipal.lst(1).Selected(frmPrincipal.lst(1).ListCount - 1) = True: DoEvents
    CreacionM NombreClase, Arr
    frmPrincipal.lst(1).AddItem "   - Edicion " + NombreClase: DoEvents
    frmPrincipal.lst(1).Selected(frmPrincipal.lst(1).ListCount - 1) = True: DoEvents
    CreacionE NombreClase, Arr
End If
End Function

Private Function Marcado(NombreClase As String) As Boolean
Dim Contador As Integer, Nombre As String
Marcado = False
For Contador = 0 To frmPrincipal.lst(0).ListCount - 1
    Nombre = UCase(frmPrincipal.lst(0).List(Contador))
    Nombre = Trim(Mid(Nombre, 1, InStr(Nombre, "(") - 1))
    If Nombre = Trim(UCase(NombreClase)) And frmPrincipal.lst(0).Selected(Contador) Then
        Marcado = True
        Exit For
    End If
Next
End Function

'OK
Public Sub CreaDD()
Dim Contador As Integer, Buffer As String
Buffer = ""
For Contador = 1 To frmPrincipal.tre(0).Nodes.Count '- 1
    If frmPrincipal.tre(0).Nodes(Contador).Children = 0 Then
        Buffer = Buffer + StrTran(frmPrincipal.tre(0).Nodes(Contador).Text, "-", " ") + "," + vbCrLf
    Else
        If Len(Buffer) > 0 Then
            Buffer = Mid(Buffer, 1, Len(Buffer) - 3) + ")" + vbCrLf
        End If
        Buffer = Buffer + vbCrLf + "create table " + frmPrincipal.tre(0).Nodes(Contador).Text + "(" + vbCrLf
    End If
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 3) + ")" + vbCrLf
Open frmPrincipal.txt(1).Text + "DD.sql" For Output As #1
Print #1, Buffer
Close #1
End Sub

'OK
Public Sub CreaINI()
Dim Contador As Integer, Buffer As String
Buffer = ""
Open frmPrincipal.txt(1).Text + frmPrincipal.txt(2).Text + ".ini" For Output As #1
Print #1, "conexion=" + cn.ConnectionString
Close #1
End Sub

'OK
Private Sub Creacion(NombreClase As String, Arr As Variant)
Dim Cadena As String, Contador As Integer, Tipo As String, Nombre As String, Buffer As String
Cadena = ""
Cadena = Cadena + "VERSION 1.0 CLASS" + vbCrLf
Cadena = Cadena + "BEGIN" + vbCrLf
Cadena = Cadena + "  MultiUse = -1  'True" + vbCrLf
Cadena = Cadena + "  Persistable = 0  'NotPersistable" + vbCrLf
Cadena = Cadena + "  DataBindingBehavior = 0  'vbNone" + vbCrLf
Cadena = Cadena + "  DataSourceBehavior  = 0  'vbNone" + vbCrLf
Cadena = Cadena + "  MTSTransactionMode  = 0  'NotAnMTSObject" + vbCrLf
Cadena = Cadena + "END" + vbCrLf
Cadena = Cadena + "Attribute VB_Name = """ + NombreClase + """" + vbCrLf
Cadena = Cadena + "Attribute VB_GlobalNameSpace = False" + vbCrLf
Cadena = Cadena + "Attribute VB_Creatable = True" + vbCrLf
Cadena = Cadena + "Attribute VB_PredeclaredId = False" + vbCrLf
Cadena = Cadena + "Attribute VB_Exposed = False" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Public MensajeError As String" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Public Arr" + NombreClase + " As Variant" + vbCrLf
Cadena = Cadena + "" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + "Public c" + Nombre + " As String" + vbCrLf
Next
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'ANADIR UN REGISTRO" + vbCrLf
Buffer = "cn As adodb.Connection, "
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Buffer = Buffer + "cl" + Nombre + " As String,"
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 1)
Cadena = Cadena + "Public Function Anadir(" + Buffer + ") As Boolean" + vbCrLf
Cadena = Cadena + "Dim rs As adodb.Recordset, Cadena As String" + vbCrLf
Cadena = Cadena + "Anadir = True" + vbCrLf
Cadena = Cadena + "On Error GoTo HELL" + vbCrLf
Cadena = Cadena + "Cadena = """"" + vbCrLf
Cadena = Cadena + "Cadena = Cadena + ""Insert into " + NombreClase + "(""" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + "Cadena = Cadena + """ + Nombre + ",""" + vbCrLf
Next
Cadena = Mid(Cadena, 1, Len(Cadena) - 4) + Chr(34) + vbCrLf
Cadena = Cadena + "Cadena = Cadena + "") values (""" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Tipo = Trim(Mid(Arr(Contador), InStr(Arr(Contador), "-") + 1))
    Tipo = Mid(Tipo, 1, 1)
    If Tipo = "C" Then
        Cadena = Cadena + "Cadena = Cadena + "" '"" + cl" + Nombre + " + ""',""" + vbCrLf
    Else
        Cadena = Cadena + "Cadena = Cadena + "" 0"" + cl" + Nombre + " + "",""" + vbCrLf
    End If
Next
Cadena = Mid(Cadena, 1, Len(Cadena) - 4) + ")" + vbCrLf
Cadena = Cadena + "Set rs = cn.Execute(Cadena)" + vbCrLf
Cadena = Cadena + "SIGUE:" + vbCrLf
Cadena = Cadena + "On Error GoTo 0" + vbCrLf
Cadena = Cadena + "Exit Function" + vbCrLf
Cadena = Cadena + "HELL:" + vbCrLf
Cadena = Cadena + "    MensajeError = Err.Description" + vbCrLf
Cadena = Cadena + "    Anadir = False" + vbCrLf
Cadena = Cadena + "    GoTo SIGUE" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "'EDITAR UN REGISTRO" + vbCrLf
Buffer = "cn As adodb.Connection, "
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Buffer = Buffer + "cl" + Nombre + " As String,"
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 1)
Cadena = Cadena + "Public Function Modificar(" + Buffer + ") As Boolean" + vbCrLf
Cadena = Cadena + "Dim rs As adodb.Recordset, Cadena As String" + vbCrLf
Cadena = Cadena + "Modificar = True" + vbCrLf
Cadena = Cadena + "On Error GoTo HELL" + vbCrLf
Cadena = Cadena + "Cadena = """"" + vbCrLf
Cadena = Cadena + "Cadena = Cadena + ""Update " + NombreClase + " set """ + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") = 0 Then
        Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
        Tipo = Trim(Mid(Arr(Contador), InStr(Arr(Contador), "-") + 1))
        Tipo = Mid(Tipo, 1, 1)
        If Tipo = "C" Then
            Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + "='"" + cl" + Nombre + " + ""', """ + vbCrLf
        Else
            Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + "= 0"" + cl" + Nombre + " + "" , """ + vbCrLf
        End If
    End If
Next
Cadena = Mid(Cadena, 1, Len(Cadena) - 5) + Chr(34) + vbCrLf
Cadena = Cadena + "Cadena = Cadena + "" where """ + vbCrLf
Cuantos = 0
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        If Cuantos > 0 Then
            Cadena = Cadena + "Cadena = Cadena + "" and """ + vbCrLf
        End If
        Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
        Tipo = Trim(Mid(Arr(Contador), InStr(Arr(Contador), "-") + 1))
        Tipo = Mid(Tipo, 1, 1)
        If Tipo = "C" Then
            Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + "='"" + cl" + Nombre + " + ""' """ + vbCrLf
        Else
            Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + "= "" + cl" + Nombre + " + ""  """ + vbCrLf
        End If
        Cuantos = Cuantos + 1
    End If
Next
Cadena = Cadena + "Set rs = cn.Execute(Cadena)" + vbCrLf
Cadena = Cadena + "SIGUE:" + vbCrLf
Cadena = Cadena + "On Error GoTo 0" + vbCrLf
Cadena = Cadena + "Exit Function" + vbCrLf
Cadena = Cadena + "HELL:" + vbCrLf
Cadena = Cadena + "    MensajeError = Err.Description" + vbCrLf
Cadena = Cadena + "    Modificar = False" + vbCrLf
Cadena = Cadena + "    GoTo SIGUE" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "'ELIMINAR" + vbCrLf
Buffer = "cn As adodb.Connection, "
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
        Buffer = Buffer + "cl" + Nombre + " As String,"
    End If
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 1)
Cadena = Cadena + "Public Function Eliminar(" + Buffer + ") As Boolean" + vbCrLf
Cadena = Cadena + "Dim rs As adodb.Recordset, Cadena As String" + vbCrLf
Cadena = Cadena + "Eliminar = True" + vbCrLf
Cadena = Cadena + "On Error GoTo HELL" + vbCrLf
Cadena = Cadena + "Cadena = """"" + vbCrLf
Cadena = Cadena + "Cadena = Cadena + ""Delete from  " + NombreClase + " where """ + vbCrLf
Cuantos = 0
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        If Cuantos > 0 Then
            Cadena = Cadena + "Cadena = Cadena + "" and """ + vbCrLf
        End If
        Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
        Tipo = Trim(Mid(Arr(Contador), InStr(Arr(Contador), "-") + 1))
        Tipo = Mid(Tipo, 1, 1)
        If Tipo = "C" Then
            Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + "='"" + cl" + Nombre + " + ""' """ + vbCrLf
        Else
            Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + "= "" + cl" + Nombre + " + ""  """ + vbCrLf
        End If
        Cuantos = Cuantos + 1
    End If
Next
Cadena = Cadena + "Set rs = cn.Execute(Cadena)" + vbCrLf
Cadena = Cadena + "SIGUE:" + vbCrLf
Cadena = Cadena + "On Error GoTo 0" + vbCrLf
Cadena = Cadena + "Exit Function" + vbCrLf
Cadena = Cadena + "HELL:" + vbCrLf
Cadena = Cadena + "    MensajeError = Err.Description" + vbCrLf
Cadena = Cadena + "    Eliminar = False" + vbCrLf
Cadena = Cadena + "    GoTo SIGUE" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'CARGAR REGISTROS" + vbCrLf
Cadena = Cadena + "Public Function Cargar(cn As adodb.Connection, Optional CantidadaDevolver As Integer, Optional CamposOrden As String) As Boolean" + vbCrLf
Cadena = Cadena + "Dim rs As adodb.Recordset, Cadena As String, Cuantos As Double" + vbCrLf
Cadena = Cadena + "Dim Ejex As Double, EjeY As Double, Columna As adodb.Field" + vbCrLf
Cadena = Cadena + "Cargar = True" + vbCrLf
Cadena = Cadena + "On Error GoTo HELL" + vbCrLf
Cadena = Cadena + "Cadena = """"" + vbCrLf
Cadena = Cadena + "Cadena = Cadena + ""select  """ + vbCrLf
Cadena = Cadena + "Cadena = Cadena + IIf(CantidadaDevolver > 0, "" top "" + CStr(CantidadaDevolver) + "" "", """")" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + ", """ + vbCrLf
Next
Cadena = Mid(Cadena, 1, Len(Cadena) - 5) + Chr(34) + vbCrLf
Cadena = Cadena + "Cadena = Cadena + "" from  " + NombreClase + " """ + vbCrLf
Cadena = Cadena + "Cadena = Cadena + IIf(Len(CamposOrden) > 0, "" order by "" + CamposOrden, """")" + vbCrLf
Cadena = Cadena + "Set rs = cn.Execute(Cadena)" + vbCrLf
Cadena = Cadena + "If Not rs.EOF Then" + vbCrLf
Cadena = Cadena + " Cuantos = 0" + vbCrLf
Cadena = Cadena + " Do While Not rs.EOF" + vbCrLf
Cadena = Cadena + "  Cuantos = Cuantos + 1" + vbCrLf
Cadena = Cadena + "  rs.MoveNext" + vbCrLf
Cadena = Cadena + " Loop" + vbCrLf
Cadena = Cadena + "rs.MoveFirst" + vbCrLf
Cadena = Cadena + " ReDim Arr" + NombreClase + "(Cuantos, rs.Fields.Count)" + vbCrLf
Cadena = Cadena + " Ejex = 0" + vbCrLf
Cadena = Cadena + " Do While Not rs.EOF" + vbCrLf
Cadena = Cadena + "  EjeY = 0" + vbCrLf
Cadena = Cadena + "  For Each Columna In rs.Fields" + vbCrLf
Cadena = Cadena + "   Arr" + NombreClase + "(Ejex, EjeY) = Columna.Value" + vbCrLf
Cadena = Cadena + "   EjeY = EjeY + 1" + vbCrLf
Cadena = Cadena + "  Next" + vbCrLf
Cadena = Cadena + "  EjeY = 0: Ejex = Ejex + 1" + vbCrLf
Cadena = Cadena + "  rs.MoveNext" + vbCrLf
Cadena = Cadena + " Loop" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "SIGUE:" + vbCrLf
Cadena = Cadena + "On Error GoTo 0" + vbCrLf
Cadena = Cadena + "Exit Function" + vbCrLf
Cadena = Cadena + "HELL:" + vbCrLf
Cadena = Cadena + "    MensajeError = Err.Description" + vbCrLf
Cadena = Cadena + "    Cargar = False" + vbCrLf
Cadena = Cadena + "    GoTo SIGUE" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'UBICAR UN REGISTRO" + vbCrLf
Buffer = "cn As adodb.Connection, "
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
        Buffer = Buffer + "cl" + Nombre + " As String,"
    End If
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 1)
Cadena = Cadena + "Public Function Ubicar(" + Buffer + ") As Boolean" + vbCrLf
Cadena = Cadena + "Dim rs As adodb.Recordset, Cadena As String" + vbCrLf
Cadena = Cadena + "Ubicar = True" + vbCrLf
Cadena = Cadena + "On Error GoTo HELL" + vbCrLf
Cadena = Cadena + "Cadena = """"" + vbCrLf
Cadena = Cadena + "Cadena = Cadena + ""select  """ + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + ", """ + vbCrLf
Next
Cadena = Mid(Cadena, 1, Len(Cadena) - 5) + Chr(34) + vbCrLf
Cadena = Cadena + "Cadena = Cadena + "" from  " + NombreClase + " where  """ + vbCrLf
Cuantos = 0
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        If Cuantos > 0 Then
            Cadena = Cadena + "Cadena = Cadena + "" and """ + vbCrLf
        End If
        Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
        Cadena = Cadena + "Cadena = Cadena + "" " + Nombre + "='"" + cl" + Nombre + " + ""' """ + vbCrLf
        Cuantos = Cuantos + 1
    End If
Next
Cadena = Cadena + "Set rs = cn.Execute(Cadena)" + vbCrLf
Cadena = Cadena + "If Not rs.EOF Then" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + " c" + Nombre + " = rs(""" + Nombre + """).Value & """"" + vbCrLf
Next
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "SIGUE:" + vbCrLf
Cadena = Cadena + "On Error GoTo 0" + vbCrLf
Cadena = Cadena + "Exit Function" + vbCrLf
Cadena = Cadena + "HELL:" + vbCrLf
Cadena = Cadena + "    MensajeError = Err.Description" + vbCrLf
Cadena = Cadena + "    Ubicar = False" + vbCrLf
Cadena = Cadena + "    GoTo SIGUE" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'SIGUIENTE REGISTRO LLAVE" + vbCrLf
Cadena = Cadena + "Public Function Siguiente(cn As adodb.Connection) As String" + vbCrLf
Cadena = Cadena + "Dim rs As adodb.Recordset, Cadena As String" + vbCrLf
Cadena = Cadena + "Siguiente = """"" + vbCrLf
Cadena = Cadena + "On Error GoTo HELL" + vbCrLf
Cadena = Cadena + "Cadena = """"" + vbCrLf
Cadena = Cadena + "Cadena = Cadena + ""Select max (""" + vbCrLf
Cuantos = 0
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        If Cuantos > 0 Then
            Cadena = Cadena + "Cadena = Cadena + "" + """ + vbCrLf
        End If
        Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
        Cadena = Cadena + "Cadena = Cadena + "" " + NombreClase + " """ + vbCrLf
        Cuantos = Cuantos + 1
    End If
Next
Cadena = Mid(Cadena, 1, Len(Cadena) - 4) + Chr(34) + vbCrLf
Cadena = Cadena + "Cadena = Cadena + "")  from " + NombreClase + " """ + vbCrLf
Cadena = Cadena + "Set rs = cn.Execute(Cadena)" + vbCrLf
Cadena = Cadena + "If Not rs.EOF Then" + vbCrLf
Cadena = Cadena + " Siguiente = rs(0)" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "SIGUE:" + vbCrLf
Cadena = Cadena + "On Error GoTo 0" + vbCrLf
Cadena = Cadena + "Exit Function" + vbCrLf
Cadena = Cadena + "HELL:" + vbCrLf
Cadena = Cadena + "    MensajeError = Err.Description" + vbCrLf
Cadena = Cadena + "    GoTo SIGUE" + vbCrLf
Cadena = Cadena + "End Function" + vbCrLf
Open frmPrincipal.txt(1).Text + NombreClase + ".cls" For Output As #1
Print #1, Cadena
Close #1
End Sub

Public Sub CreacionM(NombreClase As String, Arr As Variant)
Dim Contador As Integer, Nombre As String, Cuantos As Integer, Buffer As String
NombreClase = FCase(NombreClase)
Cadena = "" + vbCrLf
Cadena = Cadena + "VERSION 5.00" + vbCrLf
Cadena = Cadena + "Object = ""{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0""; ""MSFLXGRD.OCX""" + vbCrLf
Cadena = Cadena + "Begin VB.Form frm" + NombreClase + vbCrLf
Cadena = Cadena + "   BorderStyle     =   3  'Fixed Dialog" + vbCrLf
Cadena = Cadena + "   Caption         =   """ + NombreClase + """" + vbCrLf
Cadena = Cadena + "   ClientHeight    =   4275" + vbCrLf
Cadena = Cadena + "   ClientLeft      =   45" + vbCrLf
Cadena = Cadena + "   ClientTop       =   330" + vbCrLf
Cadena = Cadena + "   ClientWidth     =   8010" + vbCrLf
Cadena = Cadena + "   LinkTopic       =   """ + NombreClase + """" + vbCrLf
Cadena = Cadena + "   MaxButton       =   0   'False" + vbCrLf
Cadena = Cadena + "   MDIChild        =   -1  'True" + vbCrLf
Cadena = Cadena + "   MinButton       =   0   'False" + vbCrLf
Cadena = Cadena + "   ScaleHeight     =   4275" + vbCrLf
Cadena = Cadena + "   ScaleWidth      =   8010" + vbCrLf
Cadena = Cadena + "   ShowInTaskbar   =   0   'False" + vbCrLf
Cadena = Cadena + "   Begin VB.CommandButton cmd " + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Imprimir""" + vbCrLf
Cadena = Cadena + "      Height          =   375" + vbCrLf
Cadena = Cadena + "      Index           =   4" + vbCrLf
Cadena = Cadena + "      Left            =   4725" + vbCrLf
Cadena = Cadena + "      TabIndex        =   5" + vbCrLf
Cadena = Cadena + "      Top             =   3825" + vbCrLf
Cadena = Cadena + "      Width           =   1410" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin VB.ComboBox cmb " + vbCrLf
Cadena = Cadena + "      Height          =   315" + vbCrLf
Cadena = Cadena + "      Index           =   0" + vbCrLf
Cadena = Cadena + "      Left            =   4995" + vbCrLf
Cadena = Cadena + "      Style           =   2  'Dropdown List" + vbCrLf
Cadena = Cadena + "      TabIndex        =   1" + vbCrLf
Cadena = Cadena + "      Top             =   135" + vbCrLf
Cadena = Cadena + "      Width           =   2895" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin VB.CommandButton cmd " + vbCrLf
Cadena = Cadena + "      Cancel          =   -1  'True" + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Salir""" + vbCrLf
Cadena = Cadena + "      Height          =   375" + vbCrLf
Cadena = Cadena + "      Index           =   3" + vbCrLf
Cadena = Cadena + "      Left            =   6480" + vbCrLf
Cadena = Cadena + "      TabIndex        =   6" + vbCrLf
Cadena = Cadena + "      Top             =   3825" + vbCrLf
Cadena = Cadena + "      Width           =   1410" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin VB.CommandButton cmd " + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Eliminar""" + vbCrLf
Cadena = Cadena + "      Height          =   375" + vbCrLf
Cadena = Cadena + "      Index           =   2" + vbCrLf
Cadena = Cadena + "      Left            =   3195" + vbCrLf
Cadena = Cadena + "      TabIndex        =   4" + vbCrLf
Cadena = Cadena + "      Top             =   3825" + vbCrLf
Cadena = Cadena + "      Width           =   1410" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin VB.CommandButton cmd " + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Modificar""" + vbCrLf
Cadena = Cadena + "      Height          =   375" + vbCrLf
Cadena = Cadena + "      Index           =   1" + vbCrLf
Cadena = Cadena + "      Left            =   1665" + vbCrLf
Cadena = Cadena + "      TabIndex        =   3" + vbCrLf
Cadena = Cadena + "      Top             =   3825" + vbCrLf
Cadena = Cadena + "      Width           =   1410" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin VB.CommandButton cmd " + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Añadir""" + vbCrLf
Cadena = Cadena + "      Height          =   375" + vbCrLf
Cadena = Cadena + "      Index           =   0" + vbCrLf
Cadena = Cadena + "      Left            =   135" + vbCrLf
Cadena = Cadena + "      TabIndex        =   2" + vbCrLf
Cadena = Cadena + "      Top             =   3825" + vbCrLf
Cadena = Cadena + "      Width           =   1410" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin MSFlexGridLib.MSFlexGrid msf " + vbCrLf
Cadena = Cadena + "      Height          =   3120" + vbCrLf
Cadena = Cadena + "      Index           =   0" + vbCrLf
Cadena = Cadena + "      Left            =   135" + vbCrLf
Cadena = Cadena + "      TabIndex        =   0" + vbCrLf
Cadena = Cadena + "      Top             =   585" + vbCrLf
Cadena = Cadena + "      Width           =   7800" + vbCrLf
Cadena = Cadena + "      _ExtentX        =   13758" + vbCrLf
Cadena = Cadena + "      _ExtentY        =   5503" + vbCrLf
Cadena = Cadena + "      _Version        =   393216" + vbCrLf
Cadena = Cadena + "      Rows            =   20" + vbCrLf
Cadena = Cadena + "      Cols            =   20" + vbCrLf
Cadena = Cadena + "      ScrollTrack     =   -1  'True" + vbCrLf
Cadena = Cadena + "      FocusRect       =   0" + vbCrLf
Cadena = Cadena + "      GridLinesFixed  =   1" + vbCrLf
Cadena = Cadena + "      AllowUserResizing=   1" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin VB.Label lbl " + vbCrLf
Cadena = Cadena + "      AutoSize        =   -1  'True" + vbCrLf
Cadena = Cadena + "      BackStyle       =   0  'Transparent" + vbCrLf
Cadena = Cadena + "      Caption         =   ""Mostrar :""" + vbCrLf
Cadena = Cadena + "      Height          =   195" + vbCrLf
Cadena = Cadena + "      Left            =   4320" + vbCrLf
Cadena = Cadena + "      TabIndex        =   7" + vbCrLf
Cadena = Cadena + "      Top             =   180" + vbCrLf
Cadena = Cadena + "      Width           =   615" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "End" + vbCrLf
Cadena = Cadena + "Attribute VB_Name = ""frm" + NombreClase + """" + vbCrLf
Cadena = Cadena + "Attribute VB_GlobalNameSpace = False" + vbCrLf
Cadena = Cadena + "Attribute VB_Creatable = False" + vbCrLf
Cadena = Cadena + "Attribute VB_PredeclaredId = True" + vbCrLf
Cadena = Cadena + "Attribute VB_Exposed = False" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Declare Function LockWindowUpdate Lib ""User32"" (ByVal hwndLock As Integer) As Integer" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Dim Campo As String" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub Form_Load()" + vbCrLf
Cadena = Cadena + "Screen.MousePointer = vbHourglass" + vbCrLf
Cadena = Cadena + "Campo = ""1""" + vbCrLf
Cadena = Cadena + "CargaCombo" + vbCrLf
Cadena = Cadena + "CargaDatos" + vbCrLf
Buffer = "#   |"
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Buffer = Buffer + Nombre + "            |"
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 1)
Cadena = Cadena + "msf(0).FormatString = """ + Buffer + Chr(34) + vbCrLf
Cadena = Cadena + "Screen.MousePointer = vbDefault" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub CargaCombo()" + vbCrLf
Cadena = Cadena + "cmb(0).AddItem ""<TODOS>"": cmb(0).ItemData(cmb(0).ListCount - 1) = 0" + vbCrLf
Cadena = Cadena + "cmb(0).AddItem ""Los 10 Primeros"": cmb(0).ItemData(cmb(0).ListCount - 1) = 10" + vbCrLf
Cadena = Cadena + "cmb(0).AddItem ""Los 100 Primeros"": cmb(0).ItemData(cmb(0).ListCount - 1) = 100" + vbCrLf
Cadena = Cadena + "cmb(0).AddItem ""Los 200 Primeros"": cmb(0).ItemData(cmb(0).ListCount - 1) = 200" + vbCrLf
Cadena = Cadena + "cmb(0).AddItem ""Los 500 Primeros"": cmb(0).ItemData(cmb(0).ListCount - 1) = 500" + vbCrLf
Cadena = Cadena + "cmb(0).ListIndex = 1" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub cmd_Click(Index As Integer)" + vbCrLf
Cadena = Cadena + "Dim CampoLlave As String, ArrModo As Variant" + vbCrLf
Cadena = Cadena + "ArrModo = Array(""A"", ""M"", ""E"")" + vbCrLf
Cadena = Cadena + "CampoLlave = msf(0).TextMatrix(msf(0).Row, 1)" + vbCrLf
Cadena = Cadena + "Select Case Index" + vbCrLf
Cadena = Cadena + "Case 0, 1, 2" + vbCrLf
Cadena = Cadena + "    If " + NombreClase + ".Ubicar(cn, CampoLlave) Then" + vbCrLf
Cadena = Cadena + "        frmEdita" + NombreClase + ".Modo = ArrModo(Index)" + vbCrLf
Cadena = Cadena + "        frmEdita" + NombreClase + ".Show vbModal" + vbCrLf
Cadena = Cadena + "        cmb(0).ListIndex = IIf(cmb(0).ListIndex = 0, 1, cmb(0).ListIndex)" + vbCrLf
Cadena = Cadena + "        CargaDatos" + vbCrLf
Cadena = Cadena + "        msf(0).SetFocus" + vbCrLf
Cadena = Cadena + "    Else" + vbCrLf
Cadena = Cadena + "        MsgBox ""No se ubica el registro de código:"" + CampoLlave, vbOKOnly + vbCritical, ""ATENCION""" + vbCrLf
Cadena = Cadena + "        msf(0).SetFocus" + vbCrLf
Cadena = Cadena + "    End If" + vbCrLf
Cadena = Cadena + "Case 3" + vbCrLf
Cadena = Cadena + "    Unload Me" + vbCrLf
Cadena = Cadena + "Case 4" + vbCrLf
Cadena = Cadena + "    Imprime" + vbCrLf
Cadena = Cadena + "End Select" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub CargaDatos()" + vbCrLf
Cadena = Cadena + "Dim Arr As Variant, Ok As Boolean, ContaX As Double, ContaY As Double, Buffer As String, Ejex As Integer" + vbCrLf
Cadena = Cadena + "Screen.MousePointer = VBHourGlass" + vbCrLf
Cadena = Cadena + "Ejex = msf(0).Row" + vbCrLf
Cadena = Cadena + "msf(0).Rows = 1" + vbCrLf
Cadena = Cadena + "If cmb(0).ListIndex = 0 Then" + vbCrLf
Cadena = Cadena + "    Ok = " + NombreClase + ".Cargar(cn, 0, Campo)" + vbCrLf
Cadena = Cadena + "Else" + vbCrLf
Cadena = Cadena + "    Ok = " + NombreClase + ".Cargar(cn, cmb(0).ItemData(cmb(0).ListIndex), Campo)" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "If Ok Then" + vbCrLf
Cadena = Cadena + "    LockWindowUpdate msf(0).hWnd" + vbCrLf
Cadena = Cadena + "    Arr = " + NombreClase + ".Arr" + NombreClase + vbCrLf
Cadena = Cadena + "    msf(0).Cols = UBound(Arr, 2) + 1" + vbCrLf
Cadena = Cadena + "    For ContaX = 0 To UBound(Arr, 1) - 1" + vbCrLf
Cadena = Cadena + "        Buffer = CStr(ContaX + 1) + Chr(9)" + vbCrLf
Cadena = Cadena + "        For ContaY = 0 To UBound(Arr, 2) - 1" + vbCrLf
Cadena = Cadena + "            Buffer = Buffer + CStr(Arr(ContaX, ContaY) & """") + Chr(9)" + vbCrLf
Cadena = Cadena + "        Next" + vbCrLf
Cadena = Cadena + "        msf(0).AddItem Mid(Buffer, 1, Len(Buffer) - 1)" + vbCrLf
Cadena = Cadena + "    Next" + vbCrLf
Cadena = Cadena + "    On Error Resume Next" + vbCrLf
Cadena = Cadena + "    msf(0).Row = Ejex" + vbCrLf
Cadena = Cadena + "    On Error GoTo 0" + vbCrLf
Cadena = Cadena + "    LockWindowUpdate 0&" + vbCrLf
Cadena = Cadena + "    Screen.MousePointer = VBDefault" + vbCrLf
Cadena = Cadena + "Else" + vbCrLf
Cadena = Cadena + "    Screen.MousePointer = VBDefault" + vbCrLf
Cadena = Cadena + "    MsgBox ""ERROR : "" + " + NombreClase + ".MensajeError, vbOKOnly + vbCritical, ""ATENCION""" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub msf_DblClick(Index As Integer)" + vbCrLf
Cadena = Cadena + "cmd_Click 1" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub msf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)" + vbCrLf
Cadena = Cadena + "If Button = 2 Then" + vbCrLf
Cadena = Cadena + "    Campo = CStr(msf(0).Col)" + vbCrLf
Cadena = Cadena + "    CargaDatos" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub cmb_Click(Index As Integer)" + vbCrLf
Cadena = Cadena + "If Me.Visible Then" + vbCrLf
Cadena = Cadena + "    Screen.MousePointer = vbHourglass" + vbCrLf
Cadena = Cadena + "    CargaDatos" + vbCrLf
Cadena = Cadena + "    Screen.MousePointer = vbDefault" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub Imprime()" + vbCrLf
Cadena = Cadena + "Dim ContaX As Double, ContaY As Double, Archivo As String, Linea As String" + vbCrLf
Cadena = Cadena + "Screen.MousePointer = vbHourglass" + vbCrLf
Cadena = Cadena + "Archivo = App.Path + ""\t"" + Format(Now, ""yyyymmddhhmmss"") + "".tmp""" + vbCrLf
Cadena = Cadena + "Open Archivo For Output As #1" + vbCrLf
Cadena = Cadena + "For ContaX = 0 To msf(0).Rows - 1" + vbCrLf
Cadena = Cadena + "    Linea = """"" + vbCrLf
Cadena = Cadena + "    For ContaY = 0 To msf(0).Cols - 1" + vbCrLf
Cadena = Cadena + "        Linea = Linea + msf(0).TextMatrix(ContaX, ContaY) + Chr(9)" + vbCrLf
Cadena = Cadena + "    Next" + vbCrLf
Cadena = Cadena + "    Print #1, Linea" + vbCrLf
Cadena = Cadena + "Next" + vbCrLf
Cadena = Cadena + "Close #1" + vbCrLf
Cadena = Cadena + "Screen.MousePointer = vbDefault" + vbCrLf
Cadena = Cadena + "Shell ""notepad.exe "" + Archivo, vbMaximizedFocus" + vbCrLf
Cadena = Cadena + "Kill Archivo" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Open frmPrincipal.txt(1).Text + "frm" + NombreClase + ".frm" For Output As #1
Print #1, Cadena
Close #1
End Sub

Private Sub CreacionE(NombreClase As String, Arr As Variant)
Dim Contador As Integer, Nombre As String, Cuantos As Integer, Buffer As String
NombreClase = FCase(NombreClase)
Cuantos = 1300
For Contador = 0 To UBound(Arr, 1) - 1
    Cuantos = Cuantos + 400
Next
Cadena = "" + vbCrLf
Cadena = Cadena + "VERSION 5.00" + vbCrLf
Cadena = Cadena + "Begin VB.Form frmEdita" + NombreClase + vbCrLf
Cadena = Cadena + "   BorderStyle     =   3  'Fixed Dialog" + vbCrLf
Cadena = Cadena + "   Caption         =   """ + NombreClase + """" + vbCrLf
Cadena = Cadena + "   ClientHeight    =   " + CStr(Cuantos) + vbCrLf
Cadena = Cadena + "   ClientLeft      =   45" + vbCrLf
Cadena = Cadena + "   ClientTop       =   330" + vbCrLf
Cadena = Cadena + "   ClientWidth     =   6600" + vbCrLf
Cadena = Cadena + "   KeyPreview      =   -1  'True" + vbCrLf
Cadena = Cadena + "   LinkTopic       =   """ + NombreClase + """" + vbCrLf
Cadena = Cadena + "   MaxButton       =   0   'False" + vbCrLf
Cadena = Cadena + "   MinButton       =   0   'False" + vbCrLf
Cadena = Cadena + "   ScaleHeight     =   " + CStr(Cuantos) + vbCrLf
Cadena = Cadena + "   ScaleWidth      =   6600" + vbCrLf
Cadena = Cadena + "   ShowInTaskbar   =   0   'False" + vbCrLf
Cadena = Cadena + "   StartUpPosition =   2  'CenterScreen" + vbCrLf
Cadena = Cadena + "   Begin VB.CommandButton cmd " + vbCrLf
Cadena = Cadena + "      Cancel          =   -1  'True" + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Cancelar""" + vbCrLf
Cadena = Cadena + "      Height          =   420" + vbCrLf
Cadena = Cadena + "      Index           =   1" + vbCrLf
Cadena = Cadena + "      Left            =   5220" + vbCrLf
Cadena = Cadena + "      TabIndex        =   99" + vbCrLf
Cadena = Cadena + "      Top             =   " + CStr(Cuantos - 500) + vbCrLf
Cadena = Cadena + "      Width           =   1230" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cadena = Cadena + "   Begin VB.CommandButton cmd " + vbCrLf
Cadena = Cadena + "      Caption         =   ""&Aceptar""" + vbCrLf
Cadena = Cadena + "      Height          =   420" + vbCrLf
Cadena = Cadena + "      Index           =   0" + vbCrLf
Cadena = Cadena + "      Left            =   3780" + vbCrLf
Cadena = Cadena + "      TabIndex        =   98" + vbCrLf
Cadena = Cadena + "      Top             =   " + CStr(Cuantos - 500) + vbCrLf
Cadena = Cadena + "      Width           =   1230" + vbCrLf
Cadena = Cadena + "   End" + vbCrLf
Cuantos = 100
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + "   Begin VB.TextBox txt " + vbCrLf
    Cadena = Cadena + "      Height          =   330" + vbCrLf
    Cadena = Cadena + "      Index           =   " + CStr(Contador) + vbCrLf
    Cadena = Cadena + "      Left            =   2340" + vbCrLf
    Cadena = Cadena + "      TabIndex        =   " + CStr(Contador) + vbCrLf
    Cadena = Cadena + "      Top             =   " + CStr(Cuantos) + vbCrLf
    Cadena = Cadena + "      Width           =   4110" + vbCrLf
    Cadena = Cadena + "   End" + vbCrLf
    Cadena = Cadena + "   Begin VB.Label lbl " + vbCrLf
    Cadena = Cadena + "      AutoSize        =   -1  'True" + vbCrLf
    Cadena = Cadena + "      BackStyle       =   0  'Transparent" + vbCrLf
    Cadena = Cadena + "      Caption         =   """ + Nombre + ":""" + vbCrLf
    Cadena = Cadena + "      Height          =   195" + vbCrLf
    Cadena = Cadena + "      Index           =   " + CStr(Contador) + vbCrLf
    Cadena = Cadena + "      Left            =   225" + vbCrLf
    Cadena = Cadena + "      TabIndex        =   " + CStr(Contador + 100) + vbCrLf
    Cadena = Cadena + "      Top             =   " + CStr(Cuantos) + vbCrLf
    Cadena = Cadena + "      Width           =   630" + vbCrLf
    Cadena = Cadena + "   End" + vbCrLf
    Cuantos = Cuantos + 400
Next
Cadena = Cadena + "End" + vbCrLf
Cadena = Cadena + "Attribute VB_Name = ""frmEdita" + NombreClase + """" + vbCrLf
Cadena = Cadena + "Attribute VB_GlobalNameSpace = False" + vbCrLf
Cadena = Cadena + "Attribute VB_Creatable = False" + vbCrLf
Cadena = Cadena + "Attribute VB_PredeclaredId = True" + vbCrLf
Cadena = Cadena + "Attribute VB_Exposed = False" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private mModo As String" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Public Property Let Modo(vModo As String)" + vbCrLf
Cadena = Cadena + "mModo = vModo" + vbCrLf
Cadena = Cadena + "End Property" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Public Property Get Modo() As String" + vbCrLf
Cadena = Cadena + "Modo = mModo" + vbCrLf
Cadena = Cadena + "End Property" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'MODOS POSIBLES: (A)NADIR (M)ODIFICAR (E)LIMINAR" + vbCrLf
Cadena = Cadena + "Private Sub Form_Load()" + vbCrLf
Cadena = Cadena + "If Modo <> ""A"" Then" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + "    txt(" + CStr(Contador) + ").Text = " + NombreClase + ".c" + Nombre + "" + vbCrLf
Next
Cadena = Cadena + "    If Modo = ""M"" Then" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        Cadena = Cadena + "        txt(" + Str(Contador) + ").Enabled = False" + vbCrLf
    End If
Next
Cadena = Cadena + "    Else" + vbCrLf
For Contador = 0 To UBound(Arr, 1) - 1
    Nombre = FCase(Trim(Mid(Arr(Contador), 1, InStr(Arr(Contador), "-") - 1)))
    Cadena = Cadena + "    txt(" + CStr(Contador) + ").Enabled = False" + vbCrLf
Next
Cadena = Cadena + "    End If" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "'COMANDOS" + vbCrLf
Cadena = Cadena + "Private Sub cmd_Click(Index As Integer)" + vbCrLf
Cadena = Cadena + "Dim Ok As Boolean" + vbCrLf
Cadena = Cadena + "Ok = True" + vbCrLf
Cadena = Cadena + "Select Case Index" + vbCrLf
Cadena = Cadena + "Case 0" + vbCrLf
Cadena = Cadena + "    Select Case Modo" + vbCrLf
Cadena = Cadena + "    Case ""A""" + vbCrLf
Buffer = "cn, "
For Contador = 0 To UBound(Arr, 1) - 1
    Buffer = Buffer + "txt(" + CStr(Contador) + ").Text,"
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 1)
Cadena = Cadena + "        Ok = " + NombreClase + ".Anadir(" + Buffer + ")" + vbCrLf
Cadena = Cadena + "    Case ""M""" + vbCrLf
Cadena = Cadena + "        Ok = " + NombreClase + ".Modificar(" + Buffer + ")" + vbCrLf
Cadena = Cadena + "    Case ""E""" + vbCrLf
Buffer = "cn, "
For Contador = 0 To UBound(Arr, 1) - 1
    If InStr(Arr(Contador), "*") > 0 Then
        Buffer = Buffer + "txt(" + CStr(Contador) + ").Text,"
    End If
Next
Buffer = Mid(Buffer, 1, Len(Buffer) - 1)
Cadena = Cadena + "        Ok = " + NombreClase + ".Eliminar(" + Buffer + ")" + vbCrLf
Cadena = Cadena + "    End Select" + vbCrLf
Cadena = Cadena + "    If Ok Then" + vbCrLf
Cadena = Cadena + "        Unload Me" + vbCrLf
Cadena = Cadena + "    Else" + vbCrLf
Cadena = Cadena + "        MsgBox ""No se puede grabar registro: "" + " + NombreClase + ".MensajeError, vbOKOnly + vbCritical, ""ATENCION""" + vbCrLf
Cadena = Cadena + "    End If" + vbCrLf
Cadena = Cadena + "Case 1" + vbCrLf
Cadena = Cadena + "    Unload Me" + vbCrLf
Cadena = Cadena + "End Select" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub txt_GotFocus(Index As Integer)" + vbCrLf
Cadena = Cadena + "txt(Index).SelStart = 0" + vbCrLf
Cadena = Cadena + "txt(Index).SelLength = Len(txt(Index).Text)" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Cadena = Cadena + "Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)" + vbCrLf
Cadena = Cadena + "If KeyCode = 13 Or KeyCode = 40 Then" + vbCrLf
Cadena = Cadena + "    SendKeys Chr(9)" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "If KeyCode = 38 Then" + vbCrLf
Cadena = Cadena + "    SendKeys ""+"" + Chr(9)" + vbCrLf
Cadena = Cadena + "End If" + vbCrLf
Cadena = Cadena + "End Sub" + vbCrLf
Cadena = Cadena + "" + vbCrLf
Open frmPrincipal.txt(1).Text + "frmEdita" + NombreClase + ".frm" For Output As #1
Print #1, Cadena
Close #1
End Sub

