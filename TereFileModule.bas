Attribute VB_Name = "TereFileModule"
Option Explicit

'// ************************************************************
'// Modulo para el manejo de archivos y datos
'// programa para Consultorio de Estetica y cosmiatria.
'// ************************************************************
'// Fecha de inicio: 20-06-2022
'// Ultima modificación / correccion: 20-06-2022
'// ************************************************************
'// Prog.: Christian A. Del Monte - Programador
'//        Correo: creadig@gmail.com / creadig@hotmail.com
'// Rev.: Teresita Kukol - Cosmiatra
'//       Correo: kukolteresita@hotmail.com
'// ************************************************************
'// Software Name: Estetic Soft (rev 1)
'// ************************************************************

'// Datos de los diferentes pacientes
Private Type DataPaciente
    ID As Long
    FechaAlta As String * 10
    DNI As String * 10
    Apellido As String * 15
    Nombre As String * 15
    Domicilio As String * 50
    Localidad As String * 10
    Telefono1 As String * 15
    Telefono2 As String * 15
    FechaNac As String * 10
    Edad As Long
    EstadoCivil As String * 10
    Ocupacion As String * 15
    NumeroHijos As Long
End Type

'// ficha de consulta de los pacientes
Private Type FichaPaciente
    ID As Long
    FechaAlta As String * 10
    DNIpaciente As String * 10
    MotivoConsulta As String * 20
    Alergias As Boolean
        DescAlergias As String * 20
    TomaSol As Boolean
        OtroSol As String * 20
    UsaPantalla As Boolean
        OtroPantalla As String * 20
    TomaAgua As Boolean
        OtroAgua As String * 20
    Alimentacion As String * 20
    EnfermedadCronica As Boolean
        OtraEnfermedad As String * 20
    UsaMedicacion As Boolean
        OtraMedicacion As String * 20
    Diu As Boolean
    Biotipo As String * 20
    Fototipo As Long
    NumeroSesiones As Long
    Observaciones As String * 300
End Type

'// tratamientos y sesiones de los pacientes
Private Type SesionPaciente
    ID As Long
    FechaSesion As String * 10
    IDFicha As Long
    DNIpaciente As String * 10
    Tratamiento As String * 300
End Type

Public Paciente As DataPaciente
Public Ficha As FichaPaciente
Public Sesion As SesionPaciente

Public Const FilePaciente = "\data\Pacientes.pct"
Public Const FileFicha = "\data\Fichas.fch"
Public Const FileSesion = "\data\Sesiones.ssn"

'// ***************************************************************************************************
'/// Funcion para guardar los pacientes
Public Function GuardaPaciente(Dato As DataPaciente, WOptionalID As Long) As Boolean

Dim LastReg As Long

'/// abrimos el archivo de pacientes
'On Error GoTo err
Open App.Path & FilePaciente For Random As #12 Len = Len(Paciente)

'/// chequeamos por el ID de registro a guardar
If WOptionalID = 0 Or WOptionalID = -1 Then
    LastReg = LOF(12) \ Len(Paciente)
    LastReg = LastReg + 1
Else
    LastReg = WOptionalID
End If

'/// seteamos los datos del PACIENTE
Paciente = Dato
Paciente.ID = LastReg

'/// guardamos
Put #12, LastReg, Paciente
Close #12

GuardaPaciente = True
Exit Function

'/// Si hay error ------------------------------------------
'err:
'WriteErrors "Error en SaveINVData > Module INV. - " & err.Number & " - " & err.Description
'Close #12
'GuardaPaciente = False
'MsgBox "error al guardar paciente. revisar"

End Function
