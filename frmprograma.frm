VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmprograma 
   BackColor       =   &H001C1C1C&
   Caption         =   "Virtual Switch v1.0"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   7590
   ForeColor       =   &H00000000&
   Icon            =   "frmprograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin MSCommLib.MSComm USB 
      Left            =   3480
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VirtualSwitch.Panel Panel1 
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   1931
   End
   Begin VirtualSwitch.programa programa1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   11880
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu SaveAs 
         Caption         =   "&Save As..."
      End
   End
   Begin VB.Menu settings 
      Caption         =   "Settings"
   End
   Begin VB.Menu USBx 
      Caption         =   "&USB"
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu Help1 
         Caption         =   "Help"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu reloj 
      Caption         =   "----"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FRMPROGRAMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public puerto As Integer

Private Sub About_Click()
frmAbout.Show 1
End Sub

Private Sub Form_Load()
enumeradores.integrarColor
frmsettings.CargarLED "LEDS.ini"



'registro la estencion del archivo de el programa
 archivoF.CrearAsociacion App.Path & "\" & App.EXEName, _
 "vsh", "Virtual Switch" & " v1.0", App.Path & "\" & "sh.dll,0"
cargar_Driver
cargarPuerto

End Sub

Private Sub Form_Unload(Cancel As Integer)



If programa1.estado = True Then
Cancel = True
Select Case MsgBox("There is a sequence track in the program, do you want to exit the same way?", vbInformation + vbYesNo, "Advertencia")
Case (vbYes)
Cancel = False
With USB
     .Output = "1"
     .Output = "3"
     .Output = "5"
     .Output = "7"
     .Output = "9"
     Cancel = False
     End With

Case (vbNo)
Cancel = True
End Select


End If


End Sub

Private Sub New_Click()
 On Error GoTo nose
    programa1.crearSecuencia
nose:
End Sub

Private Sub Open_Click()
programa1.cmdAbrir_Click
End Sub










Private Sub Save_Click()
programa1.cmdGuardar_Click
End Sub

Private Sub SaveAs_Click()
programa1.cmdGuardarComo_Click
End Sub

Private Sub settings_Click()
'MsgBox "Todavia no esta construido este módulo de código, pero puedes observar"
frmsettings.Show 1
End Sub

Private Sub USBx_Click()
frmpuerto.Show 1
End Sub


Public Sub cargarPuerto()
 On Error GoTo nose
With USB
    .RThreshold = 1
    .InputLen = 1
    .settings = "9600" 'velocidad en baudios
    .CommPort = puerto          'numero de puerto utilizado // definalo en su arduino es el puerto donde
                           'el arduino esta conectado
                           
    .InBufferSize = 1  'Tamano del Buffer de entrada
    .InputLen = 1      'cantidad de datos a leer
    .DTREnable = False 'Deshabilitar el Threshold para TR
    .PortOpen = True   ' Abre el puerto"
End With
nose:





 
End Sub

Private Sub cargar_Driver()
Dim driv As String
On Error GoTo nose
Open App.Path & "\Drivers.hex" For Input As 1
 Do While Not EOF(1)
  Line Input #1, driv
  puerto = (driv)
 Loop
 Close #1
nose:
End Sub









