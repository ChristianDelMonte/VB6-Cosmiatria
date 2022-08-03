VERSION 5.00
Begin VB.Form Altapaciente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta de Pacientes"
   ClientHeight    =   3735
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   405
      Left            =   7800
      TabIndex        =   1
      Top             =   3000
      Width           =   1365
   End
   Begin VB.CommandButton Cmdguardar 
      Caption         =   "&Guardar"
      Height          =   405
      Left            =   6180
      TabIndex        =   0
      Top             =   3000
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   210
      ScaleHeight     =   3315
      ScaleWidth      =   9285
      TabIndex        =   2
      Top             =   210
      Width           =   9285
      Begin Proyecto1.ucText Txtdni 
         Height          =   375
         Left            =   1770
         TabIndex        =   3
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":0000
         ImgRight        =   "Altapaciente.frx":0018
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":0030
      End
      Begin Proyecto1.ucText Txtnombre 
         Height          =   345
         Left            =   1770
         TabIndex        =   4
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":0060
         ImgRight        =   "Altapaciente.frx":0078
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":0090
      End
      Begin Proyecto1.ucText Txtapellido 
         Height          =   345
         Left            =   4500
         TabIndex        =   5
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":00BC
         ImgRight        =   "Altapaciente.frx":00D4
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":00EC
      End
      Begin Proyecto1.ucText Txtdomicilio 
         Height          =   345
         Left            =   1770
         TabIndex        =   6
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":011C
         ImgRight        =   "Altapaciente.frx":0134
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":014C
      End
      Begin Proyecto1.ucText Txtlocalidad 
         Height          =   345
         Left            =   6810
         TabIndex        =   7
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":017E
         ImgRight        =   "Altapaciente.frx":0196
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":01AE
      End
      Begin Proyecto1.ucText Txtnac 
         Height          =   345
         Left            =   1770
         TabIndex        =   8
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":01E0
         ImgRight        =   "Altapaciente.frx":01F8
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":0210
      End
      Begin Proyecto1.ucText Txtestado 
         Height          =   345
         Left            =   1770
         TabIndex        =   9
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":0244
         ImgRight        =   "Altapaciente.frx":025C
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":0274
      End
      Begin Proyecto1.ucText Txtcel 
         Height          =   345
         Left            =   4500
         TabIndex        =   10
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":02AC
         ImgRight        =   "Altapaciente.frx":02C4
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":02DC
      End
      Begin Proyecto1.ucText Txtotrocel 
         Height          =   345
         Left            =   7170
         TabIndex        =   11
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":0310
         ImgRight        =   "Altapaciente.frx":0328
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":0340
      End
      Begin Proyecto1.ucText Txtocupacion 
         Height          =   345
         Left            =   4500
         TabIndex        =   12
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":0374
         ImgRight        =   "Altapaciente.frx":038C
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":03A4
      End
      Begin Proyecto1.ucText Txtedad 
         Height          =   345
         Left            =   7800
         TabIndex        =   13
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":03D6
         ImgRight        =   "Altapaciente.frx":03EE
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":0406
      End
      Begin Proyecto1.ucText Txthijos 
         Height          =   345
         Left            =   8460
         TabIndex        =   14
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
         TextConvert     =   2
         ImgLeft         =   "Altapaciente.frx":042A
         ImgRight        =   "Altapaciente.frx":0442
         RightButtonStyle=   0
         CueBanner       =   "Altapaciente.frx":045A
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Núm. Hijos:"
         Height          =   225
         Left            =   7380
         TabIndex        =   27
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Ocupación:"
         Height          =   225
         Left            =   3480
         TabIndex        =   26
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000005&
         Caption         =   "Años."
         Height          =   225
         Left            =   8430
         TabIndex        =   25
         Top             =   780
         Width           =   525
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Edad:"
         Height          =   225
         Left            =   7050
         TabIndex        =   24
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Estado Civil:"
         Height          =   225
         Left            =   270
         TabIndex        =   23
         Top             =   2040
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Teléfono:"
         Height          =   225
         Left            =   3540
         TabIndex        =   22
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Fecha de Nac.:"
         Height          =   225
         Left            =   270
         TabIndex        =   21
         Top             =   1620
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Domicilio:"
         Height          =   225
         Left            =   270
         TabIndex        =   20
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Nombre:"
         Height          =   225
         Left            =   270
         TabIndex        =   19
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "DNI:"
         Height          =   225
         Left            =   810
         TabIndex        =   18
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Otro:"
         Height          =   225
         Left            =   6570
         TabIndex        =   17
         Top             =   1620
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Apellido:"
         Height          =   225
         Left            =   3780
         TabIndex        =   16
         Top             =   780
         Width           =   645
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   "Localidad / provincia::"
         Height          =   225
         Left            =   5100
         TabIndex        =   15
         Top             =   1200
         Width           =   1635
      End
   End
End
Attribute VB_Name = "Altapaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cShadow As ClsShadow

Private Sub Cmdcancelar_Click()

Unload Me

End Sub

Private Sub Cmdguardar_Click()

'enviamos la peticion para guardar los datos
'pero antes realizamos algunas verificaciones
Dim result As Boolean

If Txtdni.Text <> "00000000" Then
    Paciente.DNI = Trim(Txtdni.Text)
Else
    MsgBox "Por favor introduzca el DNI del paciente.", vbExclamation
    Txtdni.SetFocus
    Exit Sub
End If

If LCase(Txtnombre.Text) <> "nombre" Then
    Paciente.Nombre = Trim(Txtnombre.Text)
Else
    MsgBox "Por favor introduzca el Nombre del paciente.", vbExclamation
    Txtnombre.SetFocus
    Exit Sub
End If

If LCase(Txtapellido.Text) <> "apellido" Then
    Paciente.Apellido = Trim(Txtapellido.Text)
Else
    MsgBox "Por favor introduzca el Apellido del paciente.", vbExclamation
    Txtapellido.SetFocus
    Exit Sub
End If

If Txtedad.Text <> "3" Then Paciente.Edad = CLng(Txtedad.Text) Else Paciente.Edad = 0
If LCase(Txtdomicilio.Text) <> "domicilio" Then Paciente.Domicilio = Trim(Txtdomicilio.Text) Else Paciente.Domicilio = "no especificado"
If LCase(Txtlocalidad.Text) <> "localidad o provincia" Then Paciente.Localidad = Trim(Txtlocalidad.Text) Else Paciente.Localidad = "no especificado"

Paciente.FechaNac = Trim(Txtnac.Text)

If Txtcel.Text <> "0000000000" Then
    Paciente.Telefono1 = Trim(Txtcel.Text)
Else
    MsgBox "Por favor introduzca al menos un teléfono de contacto.", vbExclamation
    Txtcel.SetFocus
    Exit Sub
End If

Paciente.Telefono2 = Trim(Txtotrocel.Text)

If LCase(Txtestado.Text) <> "estado civil" Then Paciente.EstadoCivil = Trim(Txtestado.Text) Else Paciente.EstadoCivil = "no especificado"
If LCase(Txtocupacion.Text) <> "ocupacion" Then Paciente.Ocupacion = Trim(Txtocupacion.Text) Else Paciente.Ocupacion = "no especificado"
If Txthijos.Text <> "3" Then Paciente.NumeroHijos = CLng(Txthijos.Text) Else Paciente.NumeroHijos = 0

'guardamos los datos ingresados
result = GuardaPaciente(Paciente, -1)

'limpiamos el formulario si los datos inresados fueron correctos


End Sub

Private Sub Form_Load()

    Set cShadow = New ClsShadow
    cShadow.Margin = 1
    cShadow.ShowBorders Picture1.hWnd, True
    
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

If Txtestado.Text <> "estado civil" Then
    Txtestado.ForeColor = &H80000008
Else
    Txtestado.ForeColor = &HE0E0E0
End If

End Sub

Private Sub TxtEstado_GotFocus()

Txtestado.SelStart = 0
Txtestado.SelLength = Len(Txtestado.Text)

End Sub

Private Sub TxtEstado_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Txtocupacion.SetFocus
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

If Txtocupacion.Text <> "ocupación" Then
    Txtocupacion.ForeColor = &H80000008
Else
    Txtocupacion.ForeColor = &HE0E0E0
End If

End Sub

Private Sub TxtOcupacion_GotFocus()

Txtocupacion.SelStart = 0
Txtocupacion.SelLength = Len(Txtocupacion.Text)

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
    Txtestado.SetFocus
End If

End Sub
