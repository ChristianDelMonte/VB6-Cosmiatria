VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form ConsultaPaciente 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Pacientes"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Cerrar"
      Height          =   405
      Left            =   8430
      TabIndex        =   43
      Top             =   8190
      Width           =   1155
   End
   Begin VB.Frame Frame5 
      Caption         =   "SESIONES CORRESPONDIENTES"
      Height          =   2145
      Left            =   90
      TabIndex        =   31
      Top             =   5940
      Width           =   9495
      Begin VB.CommandButton CmdPrintS 
         Caption         =   "&Imprimir ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8280
         TabIndex        =   42
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdDeleteS 
         Caption         =   "&Eliminar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   7170
         TabIndex        =   41
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdEditS 
         Caption         =   "&Modificar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   6060
         TabIndex        =   40
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdNuevoS 
         Caption         =   "&Nueva Sesión ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   4560
         TabIndex        =   39
         Top             =   1710
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1305
         Left            =   180
         TabIndex        =   38
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
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      Caption         =   "PACIENTE"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   90
      TabIndex        =   22
      Top             =   120
      Width           =   9495
      Begin VB.TextBox TxtDNI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   25
         Text            =   "00 000 000"
         Top             =   210
         Width           =   1905
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar ..."
         Enabled         =   0   'False
         Height          =   405
         Left            =   4830
         TabIndex        =   24
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton CmdNuevoP 
         Caption         =   "&Crear Nuevo ..."
         Height          =   405
         Left            =   6060
         TabIndex        =   23
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "DNI del Paciente:"
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "FICHAS DEL PACIENTE"
      Height          =   2145
      Left            =   90
      TabIndex        =   1
      Top             =   3690
      Width           =   9495
      Begin VB.CommandButton CmdPrintF 
         Caption         =   "&Imprimir ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8280
         TabIndex        =   37
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdDeleteF 
         Caption         =   "&Eliminar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   7170
         TabIndex        =   36
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdEditF 
         Caption         =   "&Modificar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   6060
         TabIndex        =   35
         Top             =   1710
         Width           =   975
      End
      Begin VB.CommandButton CmdNuevoF 
         Caption         =   "&Nueva Ficha ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   4560
         TabIndex        =   34
         Top             =   1710
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1305
         Left            =   180
         TabIndex        =   33
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      Caption         =   "DATOS PERSONALES"
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   90
      TabIndex        =   0
      Top             =   960
      Width           =   9495
      Begin VB.CommandButton CmdEditP 
         Caption         =   "&Modificar ..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   8310
         TabIndex        =   32
         Top             =   390
         Width           =   975
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FF80FF&
         Height          =   1785
         Left            =   210
         TabIndex        =   2
         Top             =   750
         Width           =   9075
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Otro:"
            Height          =   225
            Left            =   6390
            TabIndex        =   21
            Top             =   1050
            Width           =   555
         End
         Begin VB.Label LblTelefono2 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000-0000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7080
            TabIndex        =   20
            Top             =   990
            Width           =   1815
         End
         Begin VB.Label LblHijos 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8340
            TabIndex        =   19
            Top             =   1380
            Width           =   555
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Num. Hijos:"
            Height          =   225
            Left            =   7230
            TabIndex        =   18
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label LblOcupacion 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "OCUPACION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4500
            TabIndex        =   17
            Top             =   1380
            Width           =   2685
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ocupación:"
            Height          =   225
            Left            =   3420
            TabIndex        =   16
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Años."
            Height          =   225
            Left            =   8220
            TabIndex        =   15
            Top             =   270
            Width           =   675
         End
         Begin VB.Label LblEdad 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7650
            TabIndex        =   14
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Edad:"
            Height          =   225
            Left            =   6930
            TabIndex        =   13
            Top             =   270
            Width           =   675
         End
         Begin VB.Label LblEstadoCivil 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ES CIVIL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1680
            TabIndex        =   12
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civil:"
            Height          =   225
            Left            =   150
            TabIndex        =   11
            Top             =   1440
            Width           =   1425
         End
         Begin VB.Label LblTelefono1 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000-0000000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   4500
            TabIndex        =   10
            Top             =   990
            Width           =   1815
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono:"
            Height          =   225
            Left            =   3480
            TabIndex        =   9
            Top             =   1050
            Width           =   885
         End
         Begin VB.Label LblfechaNac 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00-00-2022"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1680
            TabIndex        =   8
            Top             =   990
            Width           =   1695
         End
         Begin VB.Label LblDomicilio 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DOMICILIO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1680
            TabIndex        =   7
            Top             =   600
            Width           =   7215
         End
         Begin VB.Label LblNombre 
            BackColor       =   &H00FFC0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "APELLIDO Y NOMBRE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1680
            TabIndex        =   6
            Top             =   210
            Width           =   4935
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Nac.:"
            Height          =   225
            Left            =   150
            TabIndex        =   5
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Domicilio:"
            Height          =   225
            Left            =   150
            TabIndex        =   4
            Top             =   660
            Width           =   1425
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido y Nombre:"
            Height          =   225
            Left            =   150
            TabIndex        =   3
            Top             =   270
            Width           =   1425
         End
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta:"
         Height          =   225
         Left            =   3900
         TabIndex        =   30
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20-06-2022"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         TabIndex        =   29
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DNI:"
         Height          =   225
         Left            =   1290
         TabIndex        =   28
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1890
         TabIndex        =   27
         Top             =   420
         Width           =   1815
      End
   End
End
Attribute VB_Name = "ConsultaPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()

Unload Me

End Sub

Private Sub CmdNuevoP_Click()

Altapaciente.Show , Me

End Sub

Private Sub TxtDNI_Change()

If Txtdni.Text <> "00 000 000" Then
    Txtdni.ForeColor = &H80000008
Else
    Txtdni.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtdni_GotFocus()

Txtdni.SelStart = 0
Txtdni.SelLength = Len(Txtdni.Text)

End Sub

Private Sub TxtDNI_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    'CmdBuscar.SetFocus
End If

End Sub
