VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UIVB"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIdioma 
      Alignment       =   1  'Right Justify
      Caption         =   "English"
      Height          =   240
      Left            =   7290
      TabIndex        =   30
      Top             =   45
      Width           =   1005
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Index           =   2
      Left            =   7095
      TabIndex        =   15
      ToolTipText     =   "Use este boton para salir del Software"
      Top             =   5910
      Width           =   1035
   End
   Begin TabDlg.SSTab sst 
      Height          =   5685
      Index           =   0
      Left            =   60
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   120
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   10028
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Selección de Origen de Datos"
      TabPicture(0)   =   "frmPrincipal.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tre(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmd(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmd(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Configuración del Proyecto"
      TabPicture(1)   =   "frmPrincipal.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(8)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbl(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbl(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbl(2)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lbl(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chk(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chk(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chk(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chk(3)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmd(4)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lst(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lst(0)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txt(2)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txt(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txt(3)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chk(4)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chk(5)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chk(6)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.TextBox txt 
         Height          =   345
         Index           =   0
         Left            =   2475
         TabIndex        =   29
         ToolTipText     =   "Si conoce la cadena de conexion a datos, ingresela aqui directamente"
         Top             =   480
         Width           =   4005
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   0
         Left            =   6495
         Picture         =   "frmPrincipal.frx":0D02
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Use este boton para crear una variable de conexion a Datos en forma guiada"
         Top             =   480
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Conectarse"
         Default         =   -1  'True
         Height          =   345
         Index           =   1
         Left            =   6975
         TabIndex        =   27
         ToolTipText     =   "Use este boton para intentar la conexion a Datos. "
         Top             =   480
         Width           =   1035
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crear Proyecto"
         Height          =   285
         Index           =   6
         Left            =   -74820
         TabIndex        =   1
         Top             =   495
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crear Ventanas de Mantenimiento"
         Height          =   285
         Index           =   5
         Left            =   -74820
         TabIndex        =   6
         Top             =   2070
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crear Script para generar DB (DD.sql)"
         Height          =   285
         Index           =   4
         Left            =   -74820
         TabIndex        =   7
         Top             =   2385
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   3
         Left            =   -70995
         TabIndex        =   12
         Text            =   "SISTEMA PRINCIPAL"
         ToolTipText     =   "Indique aqui el titulo que desea que aparezca en la ventana principal del proyecto cuando se este ejecutando."
         Top             =   5220
         Width           =   2745
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   1
         Left            =   -73110
         TabIndex        =   10
         ToolTipText     =   "Indique aqui la ruta donde se grabara el proyecto. "
         Top             =   4815
         Width           =   4860
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   2
         Left            =   -73110
         TabIndex        =   11
         Text            =   "Project"
         ToolTipText     =   "Indique aqui el nombre del proyecto"
         Top             =   5220
         Width           =   1125
      End
      Begin VB.ListBox lst 
         Height          =   1635
         Index           =   0
         Left            =   -74865
         Style           =   1  'Checkbox
         TabIndex        =   8
         ToolTipText     =   "Marque aqui que tablas tendran opcion de  mantenimiento"
         Top             =   3015
         Width           =   3885
      End
      Begin VB.ListBox lst 
         Height          =   3885
         Index           =   1
         Left            =   -70815
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   810
         Width           =   3885
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Proceder"
         Height          =   345
         Index           =   4
         Left            =   -67980
         TabIndex        =   13
         ToolTipText     =   "Use este boton para ejecutar el proceso"
         Top             =   5085
         Width           =   1080
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crear Clases"
         Height          =   285
         Index           =   3
         Left            =   -74820
         TabIndex        =   4
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crear Archivo INI para conexion a Datos"
         Height          =   285
         Index           =   2
         Left            =   -74820
         TabIndex        =   3
         Top             =   1125
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crear Módulo de Carga (Carga.bas)"
         Height          =   285
         Index           =   1
         Left            =   -74820
         TabIndex        =   2
         Top             =   810
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin VB.CheckBox chk 
         Caption         =   "Crear menú Principal (MDI.frm)"
         Height          =   285
         Index           =   0
         Left            =   -74820
         TabIndex        =   5
         Top             =   1755
         Value           =   1  'Checked
         Width           =   3840
      End
      Begin MSComctlLib.TreeView tre 
         Height          =   3930
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   1515
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   6932
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto MDI:"
         Height          =   195
         Index           =   5
         Left            =   -71895
         TabIndex        =   24
         Top             =   5265
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta del Proyecto:"
         Height          =   195
         Index           =   2
         Left            =   -74865
         TabIndex        =   23
         Top             =   4875
         Width           =   1320
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Proyecto:"
         Height          =   195
         Index           =   3
         Left            =   -74865
         TabIndex        =   22
         Top             =   5280
         Width           =   1530
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crear pantalla de Mantenimiento para"
         Height          =   195
         Index           =   9
         Left            =   -74820
         TabIndex        =   21
         Top             =   2790
         Width           =   2640
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOG DE EVENTOS:"
         Height          =   195
         Index           =   8
         Left            =   -70770
         TabIndex        =   20
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marque las tablas que serán incluídas en el Sistema."
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   19
         Top             =   1050
         Width           =   3720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marque los campos que serán llave para la creación de accesos a datos (Ctrl [+] expande todo, Ctrl [-] contrae)"
         Height          =   240
         Index           =   1
         Left            =   165
         TabIndex        =   18
         Top             =   1275
         Width           =   7815
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el Origen de Datos :"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   17
         Top             =   570
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   " "
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   14
      Top             =   6300
      Width           =   75
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   7
      Left            =   45
      TabIndex        =   26
      Top             =   5895
      Width           =   60
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Index           =   6
      Left            =   90
      TabIndex        =   25
      Top             =   5940
      Width           =   60
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Integer) As Integer


Private Sub Form_Load()
txt(0).Text = GetSetting("UIVB", "Principal", "Conexion", "")
Set cn = New ADODB.Connection
txt_Change 0
sst(0).TabEnabled(1) = False
tre(0).Visible = False
lbl(1).Visible = False
lbl(4).Visible = False
lbl(6).Caption = "UI para VB 2.0 (ADO) Beta release"
lbl(7).Caption = lbl(6).Caption
txt(1).Text = GetSetting("UIVB", "Principal", "Ruta", txt(1).Text)
txt(2).Text = GetSetting("UIVB", "Principal", "Nombre", txt(2).Text)
txt(3).Text = GetSetting("UIVB", "Principal", "Texto", txt(3).Text)
End Sub

Private Sub chk_Click(Index As Integer)
Dim Contador As Integer, Ok As Boolean
Ok = False
For Contador = 0 To chk.Count - 1
    chk(Contador).ForeColor = IIf(chk(Contador).Value = 1, RGB(0, 0, 0), RGB(128, 128, 128))
Next
HabilitaBoton
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
If MsgBox("¿Desea salir del Software UIVB?", vbYesNo + vbQuestion, "ATENCION") = vbYes Then
    Cancel = 0
End If
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 0
    SeleccionaConexion
Case 1
    If IntentarConexion() Then
        SaveSetting "UIVB", "Principal", "Conexion", cn.ConnectionString
        tre(0).Visible = True
        lbl(1).Visible = True
        lbl(2).Visible = True
        lbl(4).Visible = True
        DoEvents
        ArmaArbol
        sst(0).TabEnabled(1) = True
    Else
        SeleccionaConexion
    End If
Case 2
    Unload Me
Case 3
    ABuffer
Case 4
    Generacion
End Select
End Sub

Private Sub SeleccionaConexion()
Dim oDatalinks, sRetval
Set oDatalinks = CreateObject("DataLinks")
On Error Resume Next
sRetval = oDatalinks.PromptNew
On Error GoTo 0
If Not IsEmpty(sRetval) Then
 txt(0).Text = sRetval
End If
Set oDatalinks = Nothing
End Sub

Private Function IntentarConexion() As Boolean
On Error GoTo HELL
IntentarConexion = False
Screen.MousePointer = vbHourglass
Set cn = Nothing
Set cn = New ADODB.Connection
cn.ConnectionString = txt(0).Text
cn.CursorLocation = adUseClient
cn.Open
IntentarConexion = True
SIGUE:
On Error GoTo 0
Screen.MousePointer = vbDefault
Exit Function
HELL:
    Screen.MousePointer = vbDefault
    MsgBox "Error al conectarse : " + Err.Description, vbOKOnly + vbCritical, "ATENCION"
GoTo SIGUE
End Function

Private Sub ArmaArbol()
Dim Arr As Variant, Contador As Integer, Cadena As String, Tipo As String, Longitud As String
Dim rs As ADODB.Recordset, rs1 As ADODB.Recordset, Contador1 As Integer, Contador2 As Integer
Dim MiNodes As Node, Campo As String
Screen.MousePointer = vbHourglass
Set rs = cn.OpenSchema(adSchemaTables)
If Not rs.EOF Then
    rs.MoveFirst
    On Error Resume Next
    Do While Not rs.EOF
        If rs.Fields("Table_Type").Value = "TABLE" Then
            Contador1 = Contador1 + 1
            Set MiNodes = tre(0).Nodes.Add(, , Chr(Contador1 + 64), rs.Fields("Table_Name"))
            MiNodes.Checked = True
            Set rs1 = cn.Execute("select top 1 * from " + rs.Fields("Table_Name") + " order by 1")
            Contador2 = 0
            For Contador = 0 To rs1.Fields.Count - 1
                Select Case rs1.Fields(Contador2).Type
                Case 129, 200
                   Tipo = "Char"
                   Longitud = CStr(rs1.Fields(Contador2).DefinedSize)
                Case 131
                   Tipo = "Numeric"
                   Longitud = CStr(rs1.Fields(Contador2).DefinedSize) + ",2"
                Case 133
                   Tipo = "Date"
                   Longitud = CStr(rs1.Fields(Contador2).DefinedSize)
                Case 201
                   Tipo = "Memo"
                   Longitud = CStr(rs1.Fields(Contador2).DefinedSize)
                Case Else
                   Debug.Print Format(Now, "hhmmss") + " : " + rs1.Fields(Contador2).Name
                   Tipo = "???"
                   Longitud = CStr(rs1.Fields(Contador2).DefinedSize)
                End Select
                Set MiNodes = tre(0).Nodes.Add(Chr(Contador1 + 64), 4, Chr(Contador1 + 64) + Chr(Contador2), rs1.Fields(Contador2).Name + " - " + Tipo + "(" + Longitud + ")")
                If Contador = 0 Then
                    MiNodes.Checked = True
                End If
                Tipo = ""
                Contador2 = Contador2 + 1
                DoEvents
            Next
        End If
        rs.MoveNext
    Loop
    LockWindowUpdate tre(0).hWnd
    For Contador1 = 1 To tre(0).Nodes.Count - 1
        tre(0).Nodes(Contador1).EnsureVisible
    Next
    tre(0).Nodes(1).Selected = True
    LockWindowUpdate 0&
    Screen.MousePointer = vbDefault
    txt(0).Locked = True
    cmd(0).Enabled = False
    cmd(1).Enabled = False
Else
    Screen.MousePointer = vbDefault
    MsgBox "No se encuentra información de la BD : " + cn.ConnectionString, vbOKOnly + vbCritical, "ATENCION"
End If
On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting "UIVB", "Principal", "Ruta", txt(1).Text
SaveSetting "UIVB", "Principal", "Nombre", txt(2).Text
SaveSetting "UIVB", "Principal", "Texto", txt(3).Text
End Sub

Private Sub tre_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Contador As Integer
Screen.MousePointer = vbHourglass
If KeyCode = 107 And Shift = 2 Then
    LockWindowUpdate tre(0).hWnd
    For Contador = tre(0).Nodes.Count - 1 To 1 Step -1
        tre(0).Nodes(Contador).Expanded = True
    Next
    LockWindowUpdate 0&
    tre(0).Nodes(1).Selected = True
End If
If KeyCode = 109 And Shift = 2 Then
    LockWindowUpdate tre(0).hWnd
    For Contador = tre(0).Nodes.Count - 1 To 1 Step -1
        tre(0).Nodes(Contador).Expanded = False
    Next
    LockWindowUpdate 0&
    tre(0).Nodes(1).Selected = True
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub txt_Change(Index As Integer)
If Index = 0 Then
    cmd(1).Enabled = txt(0).Text <> ""
End If
HabilitaBoton
End Sub

Private Sub sst_Click(Index As Integer, PreviousTab As Integer)
Dim Contador As Integer, Cuantos As Integer, Buffer As String
Buffer = "("
If sst(0).Tab = 1 Then
    LockWindowUpdate lst(0).hWnd
    lst(0).Clear
    For Contador = tre(0).Nodes.Count - 1 To 1 Step -1
        If tre(0).Nodes(Contador).Checked And tre(0).Nodes(Contador).Children > 0 Then
            Buffer = IIf(Len(Buffer) = 1, " **SIN INDICES**", Mid(Buffer, 1, Len(Buffer) - 1) + ")")
            lst(0).AddItem tre(0).Nodes(Contador).Text + " " + Buffer
            lst(0).Selected(lst(0).ListCount - 1) = True
            Cuantos = Cuantos + 1
            Buffer = "("
        End If
        If Not tre(0).Nodes(Contador).Checked And tre(0).Nodes(Contador).Children > 0 Then
            Buffer = "("
        End If
        If tre(0).Nodes(Contador).Checked And tre(0).Nodes(Contador).Children = 0 Then
            Buffer = Buffer + Mid(tre(0).Nodes(Contador).Text, 1, InStr(tre(0).Nodes(Contador).Text, "-") - 2) + ","
        End If
    Next
    LockWindowUpdate 0&
End If
If chkIdioma.Visible Then
    lbl(9).Caption = "Crear pantalla de Mantenimiento para (" + CStr(Cuantos) + ") elementos"
Else
    lbl(9).Caption = "Maintenance forms to (" + CStr(Cuantos) + ") tables"
End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
txt(Index).SelStart = 0
txt(Index).SelLength = Len(txt(Index).Text)
End Sub

Private Sub lst_Click(Index As Integer)
Dim Contador As Integer, Cuantos As Integer
If Index = 0 Then
    Cuantos = 0
    For Contador = 0 To lst(0).ListCount - 1
        If lst(0).Selected(Contador) Then
            Cuantos = Cuantos + 1
        End If
    Next
    If chkIdioma.Visible Then
        lbl(9).Caption = "Crear pantalla de Mantenimiento para (" + CStr(Cuantos) + ") elementos"
    Else
        lbl(9).Caption = "Maintenance forms to (" + CStr(Cuantos) + ") tables"
    End If
    If Cuantos > 0 Then
        chk(5).Enabled = True
    Else
        chk(5).Enabled = False
        chk(5).Value = 0
    End If
End If
End Sub

Private Sub ABuffer()
Dim Var As String, Contador As Double, Buffer As String, Texto As String
Var = Clipboard.GetText
Var = StrTran(Var, Chr(34), "°")
Var = StrTran(Var, "°", Chr(34) + Chr(34))
Texto = "Cadena = " + Chr(34) + Chr(34) + vbCrLf + "Cadena = Cadena + "
Buffer = ""
For Contador = 1 To Len(Var)
    If Mid(Var, Contador, 1) = Chr(13) Then
        Buffer = Buffer + "Cadena = Cadena + " + Chr(34) + Texto + Chr(34) + "+ vbCrLf" + vbCrLf
        Texto = ""
        Contador = Contador + 1
    Else
        Texto = Texto + Mid(Var, Contador, 1)
    End If
Next
Clipboard.Clear
Clipboard.SetText Buffer
MsgBox Buffer
End Sub

Private Sub Generacion()
Dim Arr1 As Variant
Dim Arr2 As Variant
Arr1 = Array("Creando ruta ", "Creando Menu Principal", "Creando Módulo de Carga", "Creando Proyecto ", "Creando Archivo INI", "Creando Clases", "Creando ", "Creando Ventanas de Mantenimiento", "Proyecto finalizado")
Arr2 = Array("Creating path ", "Create MDI main form", "Creating BAS file", "Creating Project ", "Creating INI file", "Creating Classes", "Creating ", "Creating maintenance files", "Done!")
Screen.MousePointer = vbHourglass
lst(1).BackColor = RGB(216, 216, 216)
lst(1).Clear
lst(1).AddItem IIf(chkIdioma.Visible, Arr1(0), Arr2(0)) + txt(1).Text: DoEvents
CreaRuta txt(1).Text
Screen.MousePointer = vbDefault
lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
If chk(0).Value = 1 Then
    lst(1).AddItem IIf(chkIdioma.Visible, Arr1(1), Arr2(1)): DoEvents
    CreaMDI
    lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
End If
If chk(1).Value = 1 Then
    lst(1).AddItem IIf(chkIdioma.Visible, Arr1(2), Arr2(2)): DoEvents
    CreaCarga
    CreaProyecto
    lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
End If
If chk(6).Value = 1 Then
    lst(1).AddItem IIf(chkIdioma.Visible, Arr1(3), Arr2(3)) + txt(1).Text + txt(2).Text + ".vbp": DoEvents
    CreaProyecto
    lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
End If
If chk(2).Value = 1 Then
    lst(1).AddItem IIf(chkIdioma.Visible, Arr1(4), Arr2(4)): DoEvents
    CreaINI
    lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
End If
If chk(3).Value = 1 Then
    lst(1).AddItem IIf(chkIdioma.Visible, Arr1(5), Arr2(5)): DoEvents
    CreaClase
    lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
End If
If chk(4).Value = 1 Then
    lst(1).AddItem IIf(chkIdioma.Visible, Arr1(6), Arr2(6)) + txt(1).Text + "dd.sql": DoEvents
    CreaDD
    lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
End If
If chk(5).Value = 1 Then
    lst(1).AddItem IIf(chkIdioma.Visible, Arr1(7), Arr2(7)): DoEvents
    CreaMantenimiento
    lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
End If
lst(1).AddItem IIf(chkIdioma.Visible, Arr1(8), Arr2(8)): DoEvents
lst(1).Selected(lst(1).ListCount - 1) = True: DoEvents
lst(1).ListIndex = 0: DoEvents
lst(1).ListIndex = lst(1).ListCount - 1
lst(1).BackColor = RGB(255, 255, 255)
End Sub

Private Sub HabilitaBoton()
Dim Contador As Integer, Ok As Boolean
Ok = False
For Contador = 0 To chk.Count - 1
    If chk(Contador).Value = 1 Then
        Ok = True
        Exit For
    End If
Next
If Ok Then
    Ok = Len(Trim(txt(1).Text)) > 0 And Len(Trim(txt(2).Text)) > 0 And Len(Trim(txt(3).Text)) > 0
End If
cmd(4).Enabled = Ok
End Sub

Private Sub chkIdioma_Click()
chkIdioma.Visible = False
sst(0).TabCaption(0) = "Data Source"
sst(0).TabCaption(1) = "Configuration and GO"
lbl(0).Caption = "Select your Data Source:"
cmd(1).Caption = "&Connect"
lbl(4).Caption = "Check the tables to be included as class files"
lbl(1).Caption = "Check the key fields to data access.(Ctrl [+] expands all, Ctrl [-] contract)"
cmd(2).Caption = "&Quit"
chk(6).Caption = "Create project file"
chk(1).Caption = "Create bas files (carga.bas)"
chk(2).Caption = "Create ini file (with connection string)"
chk(3).Caption = "Create class files"
chk(0).Caption = "Create main MDI form"
chk(5).Caption = "Create maintenance forms"
chk(4).Caption = "Create Script to re-create data files"
lbl(9).Caption = "Create maintenance files to:"
lbl(8).Caption = "Event LOG"
lbl(2).Caption = "Project Path:"
lbl(3).Caption = "Project Name:"
lbl(5).Caption = "MDI Caption:"
cmd(4).Caption = "Go!"
lbl(6).Caption = "UI from VB 2.0 (ADO) Beta release"
lbl(7).Caption = lbl(6).Caption
End Sub

