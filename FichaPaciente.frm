VERSION 5.00
Begin VB.Form FichaPaciente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ficha de Pacientes"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "PACIENTE"
      Height          =   1035
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
         Height          =   420
         Left            =   1620
         TabIndex        =   25
         Text            =   "00000000"
         Top             =   420
         Width           =   1905
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar ..."
         Height          =   405
         Left            =   3630
         TabIndex        =   24
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Crear Nuevo ..."
         Height          =   405
         Left            =   4860
         TabIndex        =   23
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "DNI del Paciente:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   510
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "HISTORIAL"
      Height          =   2235
      Left            =   90
      TabIndex        =   1
      Top             =   4440
      Width           =   9495
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATOS"
      Height          =   2865
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   9495
      Begin VB.Frame Frame3 
         Height          =   1845
         Left            =   180
         TabIndex        =   2
         Top             =   750
         Width           =   9075
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Otro:"
            Height          =   225
            Left            =   6390
            TabIndex        =   21
            Top             =   1050
            Width           =   555
         End
         Begin VB.Label LblTelefono2 
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "387-4576331"
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
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
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
            Caption         =   "Num. Hijos:"
            Height          =   225
            Left            =   7230
            TabIndex        =   18
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label LblOcupacion 
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EMPLEADO"
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
            Caption         =   "Ocupación:"
            Height          =   225
            Left            =   3420
            TabIndex        =   16
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label Label14 
            Caption         =   "Años."
            Height          =   225
            Left            =   8220
            TabIndex        =   15
            Top             =   270
            Width           =   675
         End
         Begin VB.Label LblEdad 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "45"
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
            Caption         =   "Edad:"
            Height          =   225
            Left            =   6930
            TabIndex        =   13
            Top             =   270
            Width           =   675
         End
         Begin VB.Label LblEstadoCivil 
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CASADO"
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
            Caption         =   "Estado Civil:"
            Height          =   225
            Left            =   150
            TabIndex        =   11
            Top             =   1440
            Width           =   1425
         End
         Begin VB.Label LblTelefono1 
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "387-4576331"
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
            Caption         =   "Telefono:"
            Height          =   225
            Left            =   3480
            TabIndex        =   9
            Top             =   1050
            Width           =   885
         End
         Begin VB.Label LblfechaNac 
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "17-09-1976"
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
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HERMES QUIJADA 402 - CDAD. DEL MILAGRO"
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
            BackColor       =   &H80000010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "DEL MONTE CHRISTIAN ADRIAN"
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
            Caption         =   "Fecha de Nac.:"
            Height          =   225
            Left            =   150
            TabIndex        =   5
            Top             =   1050
            Width           =   1425
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Domicilio:"
            Height          =   225
            Left            =   150
            TabIndex        =   4
            Top             =   660
            Width           =   1425
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Apellido y Nombre:"
            Height          =   225
            Left            =   150
            TabIndex        =   3
            Top             =   270
            Width           =   1425
         End
      End
   End
End
Attribute VB_Name = "FichaPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscar_Click()

End Sub

Private Sub CmdNuevo_Click()

Altapaciente.Show

End Sub
