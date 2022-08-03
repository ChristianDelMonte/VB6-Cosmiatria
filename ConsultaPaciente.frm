VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form ConsultaPaciente 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Pacientes"
   ClientHeight    =   10620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   180
      ScaleHeight     =   2565
      ScaleWidth      =   9045
      TabIndex        =   18
      Top             =   1110
      Width           =   9045
      Begin VB.CommandButton CmdEditP 
         Caption         =   "&Modificar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   7740
         TabIndex        =   44
         Top             =   210
         Width           =   975
      End
      Begin Proyecto1.ucText ucText1 
         Height          =   375
         Left            =   1590
         TabIndex        =   19
         Top             =   270
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         InputType       =   1
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":0000
         ImgRight        =   "ConsultaPaciente.frx":0018
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":0030
      End
      Begin Proyecto1.ucText Txtnombre 
         Height          =   345
         Left            =   1590
         TabIndex        =   20
         Top             =   720
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":0060
         ImgRight        =   "ConsultaPaciente.frx":0078
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":0090
      End
      Begin Proyecto1.ucText Txtapellido 
         Height          =   345
         Left            =   4320
         TabIndex        =   21
         Top             =   720
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":00BC
         ImgRight        =   "ConsultaPaciente.frx":00D4
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":00EC
      End
      Begin Proyecto1.ucText Txtdomicilio 
         Height          =   345
         Left            =   1590
         TabIndex        =   22
         Top             =   1140
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":011C
         ImgRight        =   "ConsultaPaciente.frx":0134
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":014C
      End
      Begin Proyecto1.ucText Txtlocalidad 
         Height          =   345
         Left            =   6630
         TabIndex        =   23
         Top             =   1140
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":017E
         ImgRight        =   "ConsultaPaciente.frx":0196
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":01AE
      End
      Begin Proyecto1.ucText Txtnac 
         Height          =   345
         Left            =   1590
         TabIndex        =   24
         Top             =   1560
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         InputType       =   2
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":01E0
         ImgRight        =   "ConsultaPaciente.frx":01F8
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":0210
      End
      Begin Proyecto1.ucText Txtestado 
         Height          =   345
         Left            =   1590
         TabIndex        =   25
         Top             =   1980
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":0244
         ImgRight        =   "ConsultaPaciente.frx":025C
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":0274
      End
      Begin Proyecto1.ucText Txtcel 
         Height          =   345
         Left            =   4320
         TabIndex        =   26
         Top             =   1560
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         InputType       =   1
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":02AC
         ImgRight        =   "ConsultaPaciente.frx":02C4
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":02DC
      End
      Begin Proyecto1.ucText Txtotrocel 
         Height          =   345
         Left            =   6990
         TabIndex        =   27
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         InputType       =   1
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":0310
         ImgRight        =   "ConsultaPaciente.frx":0328
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":0340
      End
      Begin Proyecto1.ucText Txtocupacion 
         Height          =   345
         Left            =   4320
         TabIndex        =   28
         Top             =   1980
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":0374
         ImgRight        =   "ConsultaPaciente.frx":038C
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":03A4
      End
      Begin Proyecto1.ucText Txtedad 
         Height          =   345
         Left            =   7620
         TabIndex        =   29
         Top             =   720
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         InputType       =   1
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":03D6
         ImgRight        =   "ConsultaPaciente.frx":03EE
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":0406
      End
      Begin Proyecto1.ucText Txthijos 
         Height          =   345
         Left            =   8280
         TabIndex        =   30
         Top             =   1980
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         ParentBackColor =   -2147483643
         InputType       =   1
         BorderRadius    =   6
         Enabled         =   0   'False
         TextConvert     =   2
         ImgLeft         =   "ConsultaPaciente.frx":042A
         ImgRight        =   "ConsultaPaciente.frx":0442
         RightButtonStyle=   0
         CueBanner       =   "ConsultaPaciente.frx":045A
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Localidad / provincia::"
         Height          =   225
         Left            =   4920
         TabIndex        =   43
         Top             =   1200
         Width           =   1635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Apellido:"
         Height          =   225
         Left            =   3600
         TabIndex        =   42
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Otro:"
         Height          =   225
         Left            =   6390
         TabIndex        =   41
         Top             =   1620
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "DNI:"
         Height          =   225
         Left            =   630
         TabIndex        =   40
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Nombre:"
         Height          =   225
         Left            =   90
         TabIndex        =   39
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Domicilio:"
         Height          =   225
         Left            =   90
         TabIndex        =   38
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Fecha de Nac.:"
         Height          =   225
         Left            =   90
         TabIndex        =   37
         Top             =   1620
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Teléfono:"
         Height          =   225
         Left            =   3360
         TabIndex        =   36
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Estado Civil:"
         Height          =   225
         Left            =   90
         TabIndex        =   35
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Edad:"
         Height          =   225
         Left            =   6870
         TabIndex        =   34
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000005&
         Caption         =   "Años."
         Height          =   225
         Left            =   8250
         TabIndex        =   33
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Ocupación:"
         Height          =   225
         Left            =   3300
         TabIndex        =   32
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Núm. Hijos:"
         Height          =   225
         Left            =   7200
         TabIndex        =   31
         Top             =   2040
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   1470
      ScaleHeight     =   705
      ScaleWidth      =   6465
      TabIndex        =   13
      Top             =   210
      Width           =   6465
      Begin VB.CommandButton CmdNuevoP 
         Caption         =   "&Crear Nuevo ..."
         Height          =   405
         Left            =   4890
         TabIndex        =   16
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar ..."
         Enabled         =   0   'False
         Height          =   405
         Left            =   3660
         TabIndex        =   15
         Top             =   150
         Width           =   1095
      End
      Begin Proyecto1.ucText TxtDNI 
         Height          =   405
         Left            =   1590
         TabIndex        =   14
         Top             =   150
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483643
         BorderColorOnFocus=   16744703
         ParentBackColor =   16777215
         Text            =   "ConsultaPaciente.frx":047E
         InputType       =   1
         ImgLeft         =   "ConsultaPaciente.frx":04AE
         ImgRight        =   "ConsultaPaciente.frx":04C6
         RightButtonStyle=   0
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "DNI del Paciente:"
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Cerrar"
      Height          =   405
      Left            =   8430
      TabIndex        =   12
      Top             =   9570
      Width           =   1155
   End
   Begin VB.Frame Frame5 
      Caption         =   "SESIONES CORRESPONDIENTES"
      Height          =   2145
      Left            =   120
      TabIndex        =   1
      Top             =   7230
      Width           =   9495
      Begin VB.CommandButton CmdPrintS 
         Caption         =   "&Imprimir ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8280
         TabIndex        =   11
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdDeleteS 
         Caption         =   "&Eliminar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   7170
         TabIndex        =   10
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdEditS 
         Caption         =   "&Modificar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   6060
         TabIndex        =   9
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdNuevoS 
         Caption         =   "&Nueva Sesión ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   4560
         TabIndex        =   8
         Top             =   1710
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1305
         Left            =   180
         TabIndex        =   7
         Top             =   330
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2302
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "FICHAS DEL PACIENTE"
      Height          =   2145
      Left            =   240
      TabIndex        =   0
      Top             =   4860
      Width           =   9495
      Begin VB.CommandButton CmdPrintF 
         Caption         =   "&Imprimir ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8280
         TabIndex        =   6
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdDeleteF 
         Caption         =   "&Eliminar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   7170
         TabIndex        =   5
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdEditF 
         Caption         =   "&Modificar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   6060
         TabIndex        =   4
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdNuevoF 
         Caption         =   "&Nueva Ficha ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   4560
         TabIndex        =   3
         Top             =   1710
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1305
         Left            =   180
         TabIndex        =   2
         Top             =   330
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2302
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "ConsultaPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cShadow As ClsShadow
Private cShadow2 As ClsShadow

Private Sub CmdClose_Click()

Unload Me

End Sub

Private Sub CmdNuevoP_Click()

Altapaciente.Show , Me

End Sub

Private Sub Form_Load()

    Set cShadow = New ClsShadow
    cShadow.Margin = 1
    cShadow.ShowBorders Picture1.hWnd, True
    
    Set cShadow2 = New ClsShadow
    cShadow2.Margin = 1
    cShadow2.ShowBorders Picture2.hWnd, True
    
End Sub

Private Sub TxtDNI_Change()

If TxtDNI.Text <> "00 000 000" Then
    TxtDNI.ForeColor = &H80000008
Else
    TxtDNI.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtdni_GotFocus()

TxtDNI.SelStart = 0
TxtDNI.SelLength = Len(TxtDNI.Text)

End Sub

Private Sub TxtDNI_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    'CmdBuscar.SetFocus
End If

End Sub
