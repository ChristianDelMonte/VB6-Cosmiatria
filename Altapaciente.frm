VERSION 5.00
Begin VB.Form Altapaciente 
   Caption         =   "Alta de Pacientes"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtlocalidad 
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
      Height          =   345
      Left            =   6720
      TabIndex        =   6
      Text            =   "localidad o provincia"
      Top             =   1140
      Width           =   2235
   End
   Begin VB.CommandButton Cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   7800
      TabIndex        =   14
      Top             =   2580
      Width           =   1365
   End
   Begin VB.CommandButton Cmdguardar 
      Caption         =   "&Guardar"
      Height          =   405
      Left            =   6240
      TabIndex        =   13
      Top             =   2580
      Width           =   1395
   End
   Begin VB.Frame Frame3 
      Height          =   2295
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   9075
      Begin VB.TextBox Txtapellido 
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
         Height          =   345
         Left            =   4380
         TabIndex        =   3
         Text            =   "apellido"
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox Txthijos 
         Alignment       =   2  'Center
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
         Left            =   8280
         TabIndex        =   12
         Text            =   "3"
         Top             =   1800
         Width           =   555
      End
      Begin VB.TextBox Txtedad 
         Alignment       =   2  'Center
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
         Left            =   7650
         TabIndex        =   4
         Text            =   "35"
         Top             =   570
         Width           =   555
      End
      Begin VB.TextBox TxtOcupacion 
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
         Height          =   360
         Left            =   4500
         TabIndex        =   11
         Text            =   "ocupación"
         Top             =   1800
         Width           =   2625
      End
      Begin VB.TextBox Txtotrocel 
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
         Height          =   345
         Left            =   7050
         TabIndex        =   9
         Text            =   "000-0000000"
         Top             =   1410
         Width           =   1815
      End
      Begin VB.TextBox Txtcel 
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
         Height          =   345
         Left            =   4500
         TabIndex        =   8
         Text            =   "000-0000000"
         Top             =   1410
         Width           =   1845
      End
      Begin VB.TextBox TxtEstado 
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
         Height          =   345
         Left            =   1680
         TabIndex        =   10
         Text            =   "estado civil"
         Top             =   1800
         Width           =   1665
      End
      Begin VB.TextBox Txtnac 
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
         Height          =   345
         Left            =   1680
         TabIndex        =   7
         Text            =   "00-00-0000"
         Top             =   1410
         Width           =   1665
      End
      Begin VB.TextBox Txtdomicilio 
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
         Height          =   345
         Left            =   1680
         TabIndex        =   5
         Text            =   "domicilio"
         Top             =   1020
         Width           =   3075
      End
      Begin VB.TextBox Txtnombre 
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
         Height          =   345
         Left            =   1680
         TabIndex        =   2
         Text            =   "nombre"
         Top             =   630
         Width           =   1845
      End
      Begin VB.TextBox Txtdni 
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
         Height          =   345
         Left            =   1680
         TabIndex        =   1
         Text            =   "00 000 000"
         Top             =   240
         Width           =   2355
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Localidad / provincia::"
         Height          =   225
         Left            =   4950
         TabIndex        =   28
         Top             =   1110
         Width           =   1635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido:"
         Height          =   225
         Left            =   3660
         TabIndex        =   26
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Otro:"
         Height          =   225
         Left            =   6420
         TabIndex        =   25
         Top             =   1470
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DNI:"
         Height          =   225
         Left            =   690
         TabIndex        =   24
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre:"
         Height          =   225
         Left            =   150
         TabIndex        =   23
         Top             =   690
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Domicilio:"
         Height          =   225
         Left            =   150
         TabIndex        =   22
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Nac.:"
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   1470
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Teléfono:"
         Height          =   225
         Left            =   3480
         TabIndex        =   20
         Top             =   1470
         Width           =   885
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado Civil:"
         Height          =   225
         Left            =   150
         TabIndex        =   19
         Top             =   1860
         Width           =   1425
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Edad:"
         Height          =   225
         Left            =   6930
         TabIndex        =   18
         Top             =   690
         Width           =   675
      End
      Begin VB.Label Label14 
         Caption         =   "Años."
         Height          =   225
         Left            =   8220
         TabIndex        =   17
         Top             =   690
         Width           =   675
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Ocupación:"
         Height          =   225
         Left            =   3420
         TabIndex        =   16
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Núm. Hijos:"
         Height          =   225
         Left            =   7230
         TabIndex        =   15
         Top             =   1920
         Width           =   1005
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Apellido:"
      Height          =   225
      Left            =   5070
      TabIndex        =   27
      Top             =   1200
      Width           =   645
   End
End
Attribute VB_Name = "Altapaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdcancelar_Click()

Unload Me

End Sub

Private Sub Cmdguardar_Click()

'enviamos la peticion para guardar los datos
'pero antes realizamos algunas verificaciones
Dim result As Boolean

If Txtdni.Text <> "00 000 000" Then
    Paciente.DNI = Trim(Txtdni.Text)
Else
    MsgBox "Por favor introduzca el DNI del paciente.", vbExclamation
    Txtdni.SetFocus
    Exit Sub
End If

If Txtnombre.Text <> "nombre" Then
    Paciente.Nombre = Trim(Txtnombre.Text)
Else
    MsgBox "Por favor introduzca el Nombre del paciente.", vbExclamation
    Txtnombre.SetFocus
    Exit Sub
End If

If Txtapellido.Text <> "apellido" Then
    Paciente.Apellido = Trim(Txtapellido.Text)
Else
    MsgBox "Por favor introduzca el Apellido del paciente.", vbExclamation
    Txtapellido.SetFocus
    Exit Sub
End If

If Txtedad.Text <> "3" Then Paciente.Edad = CLng(Txtedad.Text) Else Paciente.Edad = 0
If Txtdomicilio.Text <> "domicilio" Then Paciente.Domicilio = Trim(Txtdomicilio.Text) Else Paciente.Domicilio = "no especificado"
If Txtlocalidad.Text <> "localidad o provincia" Then Paciente.Localidad = Trim(Txtlocalidad.Text) Else Paciente.Localidad = "no especificado"

Paciente.FechaNac = Trim(Txtnac.Text)

If Txtcel.Text <> "000-0000000" Then
    Paciente.Telefono1 = Trim(Txtcel.Text)
Else
    MsgBox "Por favor introduzca al menos un teléfono de contacto.", vbExclamation
    Txtcel.SetFocus
    Exit Sub
End If

Paciente.Telefono2 = Trim(Txtotrocel.Text)

If TxtEstado <> "estado civil" Then Paciente.EstadoCivil = Trim(TxtEstado.Text) Else Paciente.EstadoCivil = "no especificado"
If TxtOcupacion <> "ocupación" Then Paciente.Ocupacion = Trim(TxtOcupacion.Text) Else Paciente.Ocupacion = "no especificado"
If Txthijos <> "3" Then Paciente.NumeroHijos = CLng(Txthijos.Text) Else Paciente.NumeroHijos = 0

result = GuardaPaciente(Paciente, -1)

End Sub

Private Sub Txtapellido_Change()

If Txtapellido.Text <> "apellido" Then
    Txtapellido.ForeColor = &H80000008
Else
    Txtapellido.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtapellido_GotFocus()

Txtapellido.SelStart = 0
Txtapellido.SelLength = Len(Txtapellido.Text)

End Sub

Private Sub Txtapellido_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtedad.SetFocus
End If

End Sub

Private Sub Txtcel_Change()

If Txtcel.Text <> "000-0000000" Then
    Txtcel.ForeColor = &H80000008
Else
    Txtcel.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtcel_GotFocus()

Txtcel.SelStart = 0
Txtcel.SelLength = Len(Txtcel.Text)

End Sub

Private Sub Txtcel_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtotrocel.SetFocus
End If

End Sub

Private Sub Txtdni_Change()

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

Private Sub Txtdni_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtnombre.SetFocus
End If

End Sub

Private Sub Txtdomicilio_Change()

If Txtdomicilio.Text <> "domicilio" Then
    Txtdomicilio.ForeColor = &H80000008
Else
    Txtdomicilio.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtdomicilio_GotFocus()

Txtdomicilio.SelStart = 0
Txtdomicilio.SelLength = Len(Txtdomicilio.Text)

End Sub

Private Sub Txtdomicilio_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtlocalidad.SetFocus
End If

End Sub

Private Sub Txtedad_Change()

If Txtedad.Text <> "35" Then
    Txtedad.ForeColor = &H80000008
Else
    Txtedad.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtedad_GotFocus()

Txtedad.SelStart = 0
Txtedad.SelLength = Len(Txtedad.Text)

End Sub

Private Sub Txtedad_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtdomicilio.SetFocus
End If

End Sub

Private Sub TxtEstado_Change()

If TxtEstado.Text <> "estado civil" Then
    TxtEstado.ForeColor = &H80000008
Else
    TxtEstado.ForeColor = &HE0E0E0
End If

End Sub

Private Sub TxtEstado_GotFocus()

TxtEstado.SelStart = 0
TxtEstado.SelLength = Len(TxtEstado.Text)

End Sub

Private Sub TxtEstado_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtOcupacion.SetFocus
End If

End Sub

Private Sub Txthijos_Change()

If Txthijos.Text <> "3" Then
    Txthijos.ForeColor = &H80000008
Else
    Txthijos.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txthijos_GotFocus()

Txthijos.SelStart = 0
Txthijos.SelLength = Len(Txthijos.Text)

End Sub

Private Sub Txthijos_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Cmdguardar.SetFocus
End If

End Sub

Private Sub Txtlocalidad_Change()

If Txtlocalidad.Text <> "localidad o provincia" Then
    Txtlocalidad.ForeColor = &H80000008
Else
    Txtlocalidad.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtlocalidad_GotFocus()

Txtlocalidad.SelStart = 0
Txtlocalidad.SelLength = Len(Txtlocalidad.Text)

End Sub

Private Sub Txtlocalidad_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtnac.SetFocus
End If

End Sub

Private Sub Txtnac_Change()

If Txtnac.Text <> "00-00-0000" Then
    Txtnac.ForeColor = &H80000008
Else
    Txtnac.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtnac_GotFocus()

Txtnac.SelStart = 0
Txtnac.SelLength = Len(Txtnac.Text)

End Sub

Private Sub Txtnac_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtcel.SetFocus
End If

End Sub

Private Sub Txtnombre_Change()

If Txtnombre.Text <> "nombre" Then
    Txtnombre.ForeColor = &H80000008
Else
    Txtnombre.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtnombre_GotFocus()

Txtnombre.SelStart = 0
Txtnombre.SelLength = Len(Txtnombre.Text)

End Sub

Private Sub Txtnombre_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtapellido.SetFocus
End If

End Sub

Private Sub TxtOcupacion_Change()

If TxtOcupacion.Text <> "ocupación" Then
    TxtOcupacion.ForeColor = &H80000008
Else
    TxtOcupacion.ForeColor = &HE0E0E0
End If

End Sub

Private Sub TxtOcupacion_GotFocus()

TxtOcupacion.SelStart = 0
TxtOcupacion.SelLength = Len(TxtOcupacion.Text)

End Sub

Private Sub TxtOcupacion_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txthijos.SetFocus
End If

End Sub

Private Sub Txtotrocel_Change()

If Txtotrocel.Text <> "000-0000000" Then
    Txtotrocel.ForeColor = &H80000008
Else
    Txtotrocel.ForeColor = &HE0E0E0
End If

End Sub

Private Sub Txtotrocel_GotFocus()

Txtotrocel.SelStart = 0
Txtotrocel.SelLength = Len(Txtotrocel.Text)

End Sub

Private Sub Txtotrocel_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    TxtEstado.SetFocus
End If

End Sub
