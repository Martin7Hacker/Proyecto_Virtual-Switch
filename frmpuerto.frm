VERSION 5.00
Begin VB.Form frmpuerto 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "port of Exit"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF00FF&
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   290
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port - USB:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
   End
   Begin VirtualSwitch.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmpuerto.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VirtualSwitch.ChameleonBtn cmdnormal 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Normal"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmpuerto.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VirtualSwitch.ChameleonBtn cmdsalir 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   16777215
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmpuerto.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmpuerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
 cerrar
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdnormal_Click()
 Combo1.Text = 1
 almacenar_datos 'llamada al procedimiento
End Sub

Private Sub cmdnormal_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
FRMPROGRAMA.puerto = Combo1.Text

FRMPROGRAMA.cargarPuerto
Guardar_Driver
cerrar
End Sub

Private Sub cerrar()
 FRMPROGRAMA.Enabled = True
 Unload Me
End Sub

Private Sub almacenar_datos()
 FRMPROGRAMA.puerto = (Combo1.Text)
End Sub

Private Sub cargar_datos()
 Combo1.Text = FRMPROGRAMA.puerto
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub




Private Sub Combo1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
  KeyAscii = 8
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_Load()

 Me.Icon = FRMPROGRAMA.Icon
 
 cargar_datos
 ' Dar el puerto requerido
 For i = 1 To 77
 
 On Error Resume Next
 FRMPROGRAMA.USB.CommPort = i
 On Error Resume Next
 FRMPROGRAMA.USB.PortOpen = True
 On Error Resume Next
FRMPROGRAMA.USB.PortOpen = False
 If Err.Number = 0 Then
    Combo1.Clear
    Combo1.AddItem (Str(i))
    Combo1.Text = (Str(i))
 End If
 Next i
 
End Sub



Private Sub Form_Unload(Cancel As Integer)
 almacenar_datos
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmpuerto
End Sub

Public Sub Guardar_Driver()
On Error GoTo no_se
 Open App.Path & "\Drivers.hex" For Output As 1
 Dim g As Integer
 Print #1, (Combo1.Text)
 Close #1
no_se:
End Sub


