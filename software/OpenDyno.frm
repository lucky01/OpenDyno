VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Open Dyno V"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12015
   ForeColor       =   &H00000000&
   Icon            =   "OpenDyno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   12259.89
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   10800
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   56
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Km/h"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      Begin VB.Label DynoRev 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "DS-Digital"
            Size            =   56.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Vehicle Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5040
      TabIndex        =   36
      Top             =   4080
      Width           =   6735
      Begin VB.TextBox Comment 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   6495
      End
      Begin VB.TextBox Displacement 
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox Description 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Comment"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Displacement"
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      Filter          =   "Open Dyno Files *.txt"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Lambda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   4335
      Begin VB.Label LambdaValue 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "DS-Digital"
            Size            =   56.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.TextBox Status 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "DS-Digital"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   675
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "READY"
      Top             =   6720
      Width           =   3855
   End
   Begin VB.CommandButton RecordData 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10200
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Engine RPM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   12
      Top             =   2520
      Width           =   4335
      Begin VB.Label EngineRev 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "DS-Digital"
            Size            =   56.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Sensor Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   4575
      Begin VB.CheckBox Sensors 
         Alignment       =   1  'Rechts ausgerichtet
         Caption         =   "Enabled"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Undurchsichtig
         Height          =   200
         Left            =   4200
         Top             =   90
         Width           =   200
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   17
      Top             =   6120
      Width           =   4575
      Begin VB.Label RecordNo 
         Alignment       =   2  'Zentriert
         Caption         =   "Datasets: 0   Errors: 0/0/0"
         Height          =   255
         Left            =   480
         TabIndex        =   45
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Vehicle Setup"
      Height          =   2535
      Left            =   8520
      TabIndex        =   26
      Top             =   720
      Width           =   3135
      Begin VB.TextBox Clutch 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1440
         TabIndex        =   54
         Text            =   "0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox EngineRPMType 
         Height          =   315
         ItemData        =   "OpenDyno.frx":030A
         Left            =   1440
         List            =   "OpenDyno.frx":0314
         Style           =   2  'Dropdown-Liste
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton TransmissionValue 
         Caption         =   "0.00"
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox StrokeValue 
         Height          =   315
         ItemData        =   "OpenDyno.frx":0330
         Left            =   1440
         List            =   "OpenDyno.frx":033A
         Style           =   2  'Dropdown-Liste
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label ClutchLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Clutch @ RPM"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         Caption         =   "Engine RPM"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label TransmissionLabel 
         Alignment       =   1  'Rechts
         Caption         =   "Transmission"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         Caption         =   "Spark Ignition @"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Current Setup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   5040
      TabIndex        =   18
      Top             =   360
      Width           =   6735
      Begin VB.Frame Frame9 
         Caption         =   "Data Setup"
         Height          =   1575
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   3135
         Begin VB.ComboBox LossCalculation 
            Height          =   315
            ItemData        =   "OpenDyno.frx":0364
            Left            =   240
            List            =   "OpenDyno.frx":0371
            Style           =   2  'Dropdown-Liste
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton CorrectionFactor 
            Caption         =   "1.000"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox SmoothingValue 
            Alignment       =   1  'Rechts
            Height          =   285
            Left            =   240
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label13 
            Caption         =   "Loss Calculation"
            Height          =   255
            Left            =   1680
            TabIndex        =   40
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Air Correction"
            Height          =   255
            Left            =   1680
            TabIndex        =   35
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Smoothing Level"
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton ApplySetup 
         Caption         =   "Apply / ReCalc"
         Height          =   495
         Left            =   3480
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton SaveSetup 
         Caption         =   "Save Setup"
         Height          =   495
         Left            =   5160
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         Caption         =   "Drum Setup"
         Height          =   1455
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   3135
         Begin VB.Label CalibrationInertiaValue 
            Alignment       =   1  'Rechts
            Caption         =   "000.000"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "Kg/m²"
            Height          =   255
            Left            =   840
            TabIndex        =   50
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label14 
            Caption         =   "Calibration Inertia"
            Height          =   255
            Left            =   1440
            TabIndex        =   49
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "m"
            Height          =   255
            Left            =   840
            TabIndex        =   43
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label15 
            Caption         =   "Kg/m²"
            Height          =   255
            Left            =   840
            TabIndex        =   42
            Top             =   840
            Width           =   495
         End
         Begin VB.Label InertiaValue 
            Alignment       =   1  'Rechts
            Caption         =   "000.000"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Drum Inertia"
            Height          =   255
            Left            =   1440
            TabIndex        =   24
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label DiameterValue 
            Alignment       =   1  'Rechts
            Caption         =   "1000"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "Drum Diameter"
            Height          =   255
            Left            =   1440
            TabIndex        =   22
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Drum Pulses per Revolution"
            Height          =   255
            Left            =   840
            TabIndex        =   21
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label PulseValue 
            Alignment       =   1  'Rechts
            Caption         =   "0"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Data Acquisition"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   5040
      TabIndex        =   29
      Top             =   5880
      Width           =   6735
      Begin VB.CommandButton AnalyzeData 
         Caption         =   "Analyze"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton ViewData 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton LoadData 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Filename 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   48
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label7 
         Caption         =   "Filename:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   135
         Left            =   1800
         Top             =   960
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Ausgefüllt
         Height          =   135
         Left            =   5160
         Top             =   960
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents comm As MSComm
Attribute comm.VB_VarHelpID = -1

Public Temperature As Double
Public Pressure As Integer
Public Humidity As Integer
Public Correction As Double
Public CorrectionType As String
Public MaxRPM As Double
Public RPMPeak As Integer
Public MaxEngineRPM As Double
Public MaxPower As Double
Public PowerPeak As Integer
Public MaxTorque As Double
Public TorquePeak As Integer
Public MaxEnginePower As Double
Public EnginePowerPeak As Integer
Public MaxEngineTorque As Double
Public EngineTorquePeak As Integer
Public AccelTime As Double
Public Diameter As Double
Public Transmission As Double
Public Run As Integer
Public Remote As Boolean

Public DebugLog As Boolean

Dim Record2nd As Integer
Dim Speed2nd(10000) As Double
Dim Torque2nd(10000) As Double
Dim Power2nd(10000) As Double
Dim EnginePower2nd(10000) As Double
Dim EngineTorque2nd(10000) As Double
Dim EngineRPM2nd(10000) As Double
Dim DrumRPM2nd(10000) As Double
Dim DateTime2nd As String
Dim Description2nd As String
Dim Displacement2nd As String
Dim Comment2nd As String
Dim Transmission2nd As Double
Public MaxRPM2 As Double
Public MaxEngineRPM2 As Double
Public MaxPower2 As Double
Public MaxTorque2 As Double
Public MaxEnginePower2 As Double
Public MaxEngineTorque2 As Double
Public MaxSpeed2 As Double
Public MaxTorqueSpeed2 As Double
Public MaxPowerSpeed2 As Double
Public MaxEngineTorqueRPM2 As Double
Public MaxEnginePowerRPM2 As Double


Dim rcvbuffer As String
Dim DrumPeriod As Double
Dim Lambda(10000) As Double
Dim LambdaEnable As Boolean
Dim LambdaFactor(65 To 121) As Double
Dim Alpha(10000) As Double
Dim Smoothing As Integer
Dim Torque As Double
Dim Power As Double
Dim LossStart As Integer
Dim LossStart2 As Integer
Dim LossAlpha(10000) As Double
Dim LossTorque As Double
Dim LossPower As Double
Dim Speed As Double
Dim RawDyno(10000) As Long
Dim RawEngine(10000) As Long
Dim RawLambda(10000) As Long
Dim DrumRPM(10000) As Double
Dim EngineRPM(10000) As Double
Dim ClutchRPM As Integer
Dim SmoothEngineRPM(10000) As Double
Dim SmoothEngineRPM2nd(10000) As Double
Dim Record As Integer
Dim LossPercentage As Double
Dim StaticCalibration As Double
Dim DynamicCalibration As Double
Dim EnginePower As Double
Dim EngineTorque As Double

Dim AnalyzeWith As String
Dim AppendFilename As Boolean
Dim DataNo As Long
Dim DynoError As Integer
Dim RPMError As Integer
Dim LambdaError As Integer
Dim DateTime As String
Dim Stroke As Integer
Dim Pulse As Integer
Dim Period As Double
Dim Inertia As Double
Dim CalibrationInertia As Double
Dim Mode As String
Dim YellowRPM As Integer
Dim RedRPM As Integer
Public EngGraphYMin As Integer
Public EngGraphYMax As Integer
Public EngGraphYDiv As Integer
Public GraphYMin As Integer
Public GraphYMax As Integer
Public GraphYDiv As Integer
Public SpeedXMax As Integer
Public SpeedXMin As Integer
Public SpeedXDiv As Integer
Public RPMXMax As Integer
Public RPMXMin As Integer
Public RPMXDiv As Integer


Private Const Header As Integer = 21
Private Const Skip As Integer = 12

Private Const Pi As Double = 3.14159265358979
Private Const DatabaseExt = "\OpenDyno.ini"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area.
Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows.
Private Const PRF_OWNED = &H20&    ' Draw all owned windows.

Public Function ReadINI(Section As String, KeyName As String, Filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), Filename))
End Function

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Private Sub AnalyzeData_Click()
If AppendFilename = False Then
Call ShellAndWait(AnalyzeWith, 1)
Else
Call ShellAndWait(AnalyzeWith & " " & CommonDialog1.Filename, 1)
End If


End Sub

Private Sub ApplySetup_Click()

Unload ClimateCorrection
Unload DisplayResults
Unload TransmissionCalc

Stroke = (StrokeValue.ListIndex + 1) * 2
Transmission = CDbl(TransmissionValue.Caption)
On Error Resume Next
Smoothing = CInt(SmoothingValue.Text)
If Err Then
    Smoothing = CInt(ReadINI("Data", "Smoothing", App.Path & DatabaseExt))
    SmoothingValue.Text = Smoothing
End If
On Error Resume Next
ClutchRPM = CInt(Clutch.Text)
If Err Then
    Clutch.Text = 0
    ClutchRPM = 0
End If

If Record <> 0 Then
        Call AlphaCalc(Smoothing + 1)
        Call GetPeak
        Call DisplayResult
End If

End Sub


Private Sub EngineRPMType_Click()
If EngineRPMType.ListIndex = 0 Then
Clutch.Enabled = False
TransmissionValue.Enabled = False
ClutchLabel.Enabled = False
TransmissionLabel.Enabled = False
Else
Clutch.Enabled = True
TransmissionValue.Enabled = True
ClutchLabel.Enabled = True
TransmissionLabel.Enabled = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call SaveSetup_Click

If DebugLog = True Then Close #2

Unload ClimateCorrection
Unload DisplayResults
Unload PrinterSheet
Unload TransmissionCalc
Unload ComparePrinterSheet
End Sub

Private Sub ViewData_Click()
Remote = False
If Mode = "record" Then SwitchMode ("save") Else SwitchMode ("standby")

If Record <> 0 Then
    DisplayResult
End If
End Sub


Private Sub Sensors_Click()
If Sensors.Value = vbChecked And Mode <> "record" Then SwitchMode ("sensors")
If Sensors.Value = vbUnchecked And Mode = "sensors" Then SwitchMode ("standby")
End Sub

'---OnComm is an event-------
Private Sub comm_OnComm()
    Dim temp As String

    temp = comm.Input

    If DebugLog = True Then
     Print #2, temp;
    End If
    
    rcvbuffer = rcvbuffer & temp
If InStr(rcvbuffer, Chr(10)) <> 0 Then
    Marker = InStr(rcvbuffer, Chr(10))
    rcvtxt = Left$(rcvbuffer, Marker + 1)
    rcvbuffer = Mid$(rcvbuffer, Marker + 1)
    If InStr(rcvtxt, "DATA") <> 0 Then Status.Text = "DATA"
    If InStr(rcvtxt, "STOP") <> 0 Then Status.Text = "STOP"
    If InStr(rcvtxt, "DIAG") <> 0 Then Status.Text = "DIAG"
    If InStr(rcvtxt, "BT1") <> 0 Then Call Button1 'white
    If InStr(rcvtxt, "BT2") <> 0 Then Call Button2 'green
    If InStr(rcvtxt, "BT3") <> 0 Then Call Button3 'red
    If InStr(rcvtxt, ";") And Len(rcvtxt) > 5 Then
        Value = Split(rcvtxt, ";", 4)
        On Error Resume Next
        RawDyno(DataNo) = Value(0)
        On Error Resume Next
        RawEngine(DataNo) = Value(1)
        On Error Resume Next
        RawLambda(DataNo) = Value(2)
' Error Correction
        If DataNo > 0 And Mode = "record" Then
            If RawDyno(DataNo - 1) <> 0 Then
                If RawDyno(DataNo) / RawDyno(DataNo - 1) <= 0.5 Then
                    RawDyno(DataNo) = RawDyno(DataNo - 1)
                    DynoError = DynoError + 1
                End If
                If RawDyno(DataNo) / RawDyno(DataNo - 1) >= 2 Then
                    RawDyno(DataNo) = RawDyno(DataNo - 1)
                    DynoError = DynoError + 1
                End If
            Else
                RawDyno(DataNo) = RawDyno(DataNo - 1)
                DynoError = DynoError + 1
            End If
            If RawEngine(DataNo - 1) <> 0 Then
                If RawEngine(DataNo) / RawEngine(DataNo - 1) <= 0.5 Then
                    RawEngine(DataNo) = RawEngine(DataNo - 1)
                    RPMError = RPMError + 1
                End If
                If RawEngine(DataNo) / RawEngine(DataNo - 1) >= 2 Then
                    RawEngine(DataNo) = RawEngine(DataNo - 1)
                    RPMError = RPMError + 1
                End If
            End If
            If RawLambda(DataNo - 1) <> 0 Then
                If RawLambda(DataNo) / RawLambda(DataNo - 1) <= 0.95 Then
                    RawLambda(DataNo) = RawLambda(DataNo - 1)
                    LambdaError = LambdaError + 1
                End If
                If RawLambda(DataNo) / RawLambda(DataNo - 1) >= 1.05 Then
                    RawLambda(DataNo) = RawLambda(DataNo - 1)
                    LambdaError = LambdaError + 1
                End If
            End If
        End If
        
        DisplaySensor
        If Mode = "record" Then DataNo = DataNo + 1
        If DataNo = 9999 Then SwitchMode "save"
    End If
End If
        
End Sub

Public Sub Button2() 'green
Remote = True
Select Case Mode
    Case "standby"
        SwitchMode ("record")
    Case "record"
    Case "save"
        Unload DisplayResults
        CommonDialog1.Filename = CommonDialog1.InitDir & "\dyno" & CStr(Format(Main.Run, "000")) & "." & CommonDialog1.DefaultExt
        Call SaveData
        SwitchMode ("standby")
    Case "sensors"
        SwitchMode ("record")
End Select

End Sub

Public Sub Button3() 'red
Remote = True
Select Case Mode
    Case "standby"
        SwitchMode ("sensors")
    Case "record"
        If DataNo > Pulse Then
            SwitchMode ("save")
        Else
            SwitchMode ("standby")
        End If
    Case "save"
        If DisplayResults.DataSelection.ListIndex = 2 Then
            DisplayResults.DataSelection.ListIndex = 0
        Else
            DisplayResults.DataSelection.ListIndex = DisplayResults.DataSelection.ListIndex + 1
        End If
        Call Main.DrawChart
    
    Case "sensors"
        SwitchMode ("standby")
End Select

End Sub
Public Sub Button1() 'white
Remote = True
Select Case Mode
    Case "standby"
        If Record <> 0 Then
            DisplayResult
        End If
    Case "record"
        SwitchMode ("save")
        If Record <> 0 Then
            DisplayResult
        End If
    Case "save"
        
        Call Main.PrintChart
        CommonDialog1.Filename = CommonDialog1.InitDir & "\dyno" & CStr(Format(Main.Run, "000")) & "." & CommonDialog1.DefaultExt
        Remote = True
        Call SaveData
        Unload DisplayResults
        SwitchMode ("standby")
    Case "sensors"
        SwitchMode ("standby")
        If Record <> 0 Then
            DisplayResult
        End If
End Select

End Sub


Public Sub DataCalc(i As Integer)


Torque = ((Inertia + CalibrationInertia) * Alpha(i)) ' Torque in Nm

Torque = Torque + StaticCalibration + (DynamicCalibration / DrumPeriod)

Power = ((Torque * 2 * Pi * (1 / DrumPeriod)) / 1000) * Correction 'Power in Kilowatts

Speed = Pi * (Diameter / 1000) * (DrumRPM(i) * 60) 'Speed in km/h


If i < Skip Then
    MaxRPM = 0
End If

If DrumRPM(i) > MaxRPM Then
   MaxRPM = DrumRPM(i)
   RPMPeak = i
End If

If Torque >= 0 Then
    If i > RPMPeak Then
    If i = LossStart + 1 Then LossStart = i
    Else
    LossStart = i
    End If
    LossTorque = (Inertia + CalibrationInertia) * LossAlpha(i) ' Torque in Nm
    LossTorque = LossTorque + StaticCalibration + (DynamicCalibration / DrumPeriod)
    LossPower = (LossTorque * 2 * Pi * (1 / DrumPeriod)) / 1000 'Power in Kilowatts
    
    
    If EngineRPMType.ListIndex = 0 Then
        If EngineRPM(i) <> 0 Then
            TransmissionFactor = (EngineRPM(i) / DrumRPM(i))
        Else
            TransmissionFactor = Transmission
            SmoothEngineRPM(i) = DrumRPM(i) * Transmission
        End If
    Else
    If DrumRPM(i) * Transmission > ClutchRPM Then
    TransmissionFactor = Transmission
    Else
    TransmissionFactor = ClutchRPM / DrumRPM(i)
    End If
    End If
    
Select Case LossCalculation.ListIndex
 Case 0
    EngineTorque = (Torque - LossTorque) / TransmissionFactor
 Case 1
    EngineTorque = (Torque * (1 + (LossPercentage / 100))) / TransmissionFactor
 Case 2
    EngineTorque = Torque / TransmissionFactor
End Select
    EnginePower = ((EngineTorque * 2 * Pi * ((1 / DrumPeriod) * TransmissionFactor)) / 1000) * Correction 'Power in Kilowatt
    
Else
    LossTorque = 0
    LossPower = 0
End If

End Sub



Public Function START() As String
Set comm = New MSComm
            
    

    
            '-------READING COMM PARAMETERS
            
            comm.Settings = ReadINI("General", "BitRate", App.Path & DatabaseExt) & ",n,8,1"
            comm.Handshaking = 1
            comm.CommPort = CInt(ReadINI("General", "ComPort", App.Path & DatabaseExt))
            comm.RThreshold = 1
            comm.InBufferSize = 2048
            comm.OutBufferSize = 2048
            On Error Resume Next
            comm.PortOpen = True
            If comm.PortOpen = False Then START = "NoComm"
            '-----------------------------------
            
End Function


Private Sub DisplaySensor()

If RawDyno(DataNo) > 0 And Sensors.Value = vbChecked Then
  If (60 / (RawDyno(DataNo) * Period * Pulse)) < YellowRPM Then
    DynoRev.ForeColor = &HFF00&
  End If
  If (60 / (RawDyno(DataNo) * Period * Pulse)) >= YellowRPM Then
    DynoRev.ForeColor = &H80FF&
  End If
  If (60 / (RawDyno(DataNo) * Period * Pulse)) >= RedRPM Then
    DynoRev.ForeColor = &HFF&
  End If
  DynoRev.Caption = Format(Pi * (Diameter / (Pulse * 1000)) * 60 * (60 / (RawDyno(DataNo) * Period)), "0.0")
Else
  DynoRev.Caption = 0
  DynoRev.ForeColor = &HFF00&
End If

If RawEngine(DataNo) > 0 And Sensors.Value = vbChecked Then EngineRev.Caption = Round((60 / (RawEngine(DataNo) * Period)) / (Stroke / 2)) Else EngineRev.Caption = 0

If RawLambda(DataNo) > 0 And Sensors.Value = vbChecked And LambdaEnable = True Then
    i = 65
    Do While RawLambda(DataNo) <= LambdaFactor(i)
        i = i + 1
    Loop
    LambdaValue.Caption = Format(i / 100, "0.00")
Else
    LambdaValue.Caption = "0.00"
End If
RecordNo.Caption = "Datasets: " & DataNo & "   Errors: " & DynoError & "/" & RPMError & "/" & LambdaError
End Sub



Public Sub SwitchMode(temp As String)

Mode = temp
Select Case Mode
    Case "standby"
        Shape3.Visible = True
        Shape1.FillColor = &HFF00&
        DynoRev.Caption = "0"
        EngineRev.Caption = "0"
        LambdaValue.Caption = "0.00"
        DynoRev.ForeColor = &HFF00&
        comm.Output = "0"
        Sensors.Value = vbUnchecked
        RecordData.Caption = "Start"
    Case "sensors"
        Shape3.Visible = True
        Shape1.FillColor = &HFF00&
        comm.Output = "2"
        RecordData.Caption = "Start"
        rcvbuffer = ""
        Sensors.Value = vbChecked
    Case "record"
        Shape1.FillColor = &HFF&
        Shape3.Visible = False
        Unload ClimateCorrection
        Unload DisplayResults
        Unload TransmissionCalc
        LoadSetup
        Erase RawDyno
        Erase RawEngine
        Erase RawLambda
        DataNo = 0
        DynoError = 0
        RPMError = 0
        LambdaError = 0
        Sensors.Value = vbChecked
        RecordData.Caption = "Stop"
        comm.Output = "1"
        rcvbuffer = ""
    Case "save"
        Shape3.Visible = False
        Shape1.FillColor = &HFF00&
        DynoRev.Caption = "0"
        EngineRev.Caption = "0"
        LambdaValue.Caption = "0.00"
        DynoRev.ForeColor = &HFF00&
        DateTime = Format(Now, "Short Date") & " " & Format(Time, "Short Time")
        comm.Output = "0"
        Sensors.Value = vbUnchecked
        RecordData.Caption = "Start"

        If DataNo > Pulse Then
            Call DrumRPMCalc
            Call EngineRPMCalc
            Call LambdaCalc
            Call AlphaCalc(Smoothing + 1)
            Call GetPeak
            Call DisplayResult
        End If
        
End Select
    
   
End Sub
    

Private Sub CorrectionFactor_Click()
ClimateCorrection.Temperature.Text = Temperature
ClimateCorrection.Pressure.Text = Pressure
ClimateCorrection.Humidity.Text = Humidity
ClimateCorrection.Show
ClimateCorrection.CorrectionValue.Text = CorrectionFactor.Caption
End Sub

Private Sub RecordData_Click()
Remote = False
If Mode = "record" Then
    If DataNo > Pulse Then SwitchMode ("save") Else SwitchMode ("sensors")
Else
    SwitchMode ("record")
End If
End Sub

Private Sub DisplayResult()
Mode = "save"

With DisplayResults

.WSpeed.Caption = Format(Pi * (Diameter / 1000) * (MaxRPM * 60), "0.00")
.AccelTime.Caption = Format(AccelTime, "0.00")
.WTorque.Caption = Format(MaxTorque, "0.00")
.WTorqueSpeed.Caption = Format(Pi * (Diameter / 1000) * (DrumRPM(TorquePeak) * 60), "0.00")
.WPower.Caption = Format(MaxPower, "0.00") & " (" & Format(MaxPower * 1.359622, "0.00") & "PS)"
.WPowerSpeed.Caption = Format(Pi * (Diameter / 1000) * (DrumRPM(PowerPeak) * 60), "0.00")

If MaxEngineRPM = 0 Then
.ERPM.Caption = Format(MaxRPM * Transmission, "0")
Else
.ERPM.Caption = Format(MaxEngineRPM, "0")
End If

.ETorque.Caption = Format(MaxEngineTorque, "0.00")
If EngineRPM(EngineTorquePeak) = 0 Then
.ETorqueRPM.Caption = Format((DrumRPM(EngineTorquePeak) * Transmission), "0")
.Note.Visible = True
Else
.ETorqueRPM.Caption = Format((EngineRPM(EngineTorquePeak)), "0")
.Note.Visible = False
End If

.EPower.Caption = Format(MaxEnginePower, "0.00") & " (" & Format(MaxEnginePower * 1.359622, "0.00") & "PS)"
If EngineRPM(EnginePowerPeak) = 0 Then
.EPowerRPM.Caption = Format((DrumRPM(EnginePowerPeak) * Transmission), "0")
.Note.Visible = True
Else
.EPowerRPM.Caption = Format(EngineRPM(EnginePowerPeak), "0")
.Note.Visible = False
End If

End With


DisplayResults.DataSelection.ListIndex = 1
DisplayResults.CompareSelection.ListIndex = 0
Call DrawChart
DisplayResults.ZoomXMin.Value = SpeedXMin / 5
DisplayResults.ZoomXMax.Value = SpeedXMax / 5
DisplayResults.ZoomYMax.Value = EngGraphYMax * -1
DisplayResults.Show

End Sub

Public Sub DrawChart()

With DisplayResults.MSChart1

    .RowCount = 0
    .ColumnCount = 0
    .RandomFill = False
    .ColumnCount = 12
    
    '.Plot.SeriesCollection(1).SeriesMarker.Auto = False
    '.Plot.SeriesCollection(1).SeriesMarker.Show = False
    
    '.Plot.SeriesCollection(1).DataPoints.Item(-1).Marker.Visible = False

    
    .Plot.SeriesCollection(7).Pen.VtColor.Set 0, 255, 0
    .Plot.SeriesCollection(9).Pen.VtColor.Set 255, 165, 0
    .Plot.SeriesCollection(11).Pen.VtColor.Set 255, 255, 0
    .Plot.SeriesCollection(5).Position.Hidden = True
    .Plot.SeriesCollection(7).Position.Hidden = True
    .Plot.SeriesCollection(9).Position.Hidden = True
    .Plot.SeriesCollection(11).Position.Hidden = True

    .chartType = VtChChartType2dXY

    .Plot.AutoLayout = False
    .Plot.UniformAxis = False
    .Plot.WidthToHeightRatio = 1
    .Plot.SeriesCollection(1).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(3).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(1).Pen.Width = 25
    .Plot.SeriesCollection(3).Pen.Width = 25
    
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
    .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    
    'title area

    .Title.VtFont.Style = VtFontStyleBold
    .Title.VtFont.Name = "Arial"
    .Title.VtFont.Size = 14
    
    
    'footnote area

    '.FootnoteText = "This is a foot note"
    '.Footnote.VtFont.VtColor.Set 25, 150, 200
    
    '.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"
    '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = "Torque(Nm)"
    '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY2).AxisTitle.Text = "Power(kW)"
    
    .ShowLegend = True

    '.Column = 1
    '.ColumnLabel = "Torque(Nm)"
    '.Column = 3
    '.ColumnLabel = "Power(kW)"



Select Case DisplayResults.DataSelection.ListIndex

Case 0
    .RowCount = Record
   
    '.ColumnCount = 4
    
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    If Record2nd > Record Then .RowCount = Record2nd
    
    '.ColumnCount = 8
    .Plot.SeriesCollection(5).Position.Hidden = True
    .Plot.SeriesCollection(7).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(9).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(7).Pen.Width = 25
    .Plot.SeriesCollection(9).Pen.Width = 25
    .Column = 7
    .ColumnLabel = "2nd Torque(Nm)"
    .Column = 9
    .ColumnLabel = "2nd Power(kW)"
    End If
    
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = GraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = GraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = GraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = SpeedXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = SpeedXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = SpeedXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"

    For i = 1 To Record
        
    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    .Row = i
    .Column = 1
    .Data = Speed
    .Column = 2
    .Data = Torque
    .Column = 3
    .Data = Speed
    .Column = 4
    .Data = Power
    
    Next i
    
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    For i = 1 To Record2nd
    .Row = i
    .Column = 7
    .Data = Speed2nd(i)
    .Column = 8
    .Data = Torque2nd(i)
    .Column = 9
    .Data = Speed2nd(i)
    .Column = 10
    .Data = Power2nd(i)
    Next i
    End If
    
       
Case 1
    
    .RowCount = LossStart
    
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    If LossStart2 > LossStart Then .RowCount = LossStart2
    .Plot.SeriesCollection(7).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(7).Pen.Width = 25
    .Plot.SeriesCollection(9).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(9).Pen.Width = 25
    .Plot.SeriesCollection(11).Position.Hidden = DisplayResults.Check_RPM.Value
    .Plot.SeriesCollection(11).Pen.Width = 25
    
    .Column = 7
    .ColumnLabel = "2nd Torque(Nm)"
    .Column = 9
    .ColumnLabel = "2nd Power(kW)"
    .Column = 11
    .ColumnLabel = "2nd RPM (1/min x 1000)"
    End If
    
    .Plot.SeriesCollection(5).Position.Hidden = DisplayResults.Check_RPM.Value
    .Plot.SeriesCollection(5).Pen.Width = 25
    
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    .Column = 5
    .ColumnLabel = "RPM (1/min x 1000)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = EngGraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = EngGraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = EngGraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = SpeedXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = SpeedXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = SpeedXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"
    
    Call EngineRPMSmooth(Smoothing + 1)
    
    For i = 1 To LossStart
        
    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    .Row = i
    .Column = 1
    .Data = Speed
    .Column = 2
    .Data = EngineTorque
    .Column = 3
    .Data = Speed
    .Column = 4
    .Data = EnginePower
    .Column = 5
    .Data = Speed
    .Column = 6
    
    If EngineRPMType.ListIndex = 0 Then
    .Data = SmoothEngineRPM(i) / 1000
    Else
    If DrumRPM(i) * Transmission > ClutchRPM Then
    .Data = DrumRPM(i) * Transmission / 1000
    Else
    .Data = ClutchRPM / 1000
    End If
    End If
    
    Next i
    
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    For i = 1 To LossStart2
    .Row = i
    .Column = 7
    .Data = Speed2nd(i)
    .Column = 8
    .Data = EngineTorque2nd(i)
    .Column = 9
    .Data = Speed2nd(i)
    .Column = 10
    .Data = EnginePower2nd(i)
    .Column = 11
    .Data = Speed2nd(i)
    .Column = 12
    If EngineRPMType.ListIndex = 0 Then
     .Data = SmoothEngineRPM2nd(i) / 1000
    Else
        If DrumRPM2nd(i) * Transmission2nd > ClutchRPM Then
            .Data = DrumRPM2nd(i) * Transmission2nd / 1000
        Else
            .Data = ClutchRPM / 1000
        End If
    End If
    Next i
    End If
    
    
Case 2
    .RowCount = LossStart
    
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    If LossStart2 > LossStart Then .RowCount = LossStart2
    .Plot.SeriesCollection(7).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(7).Pen.Width = 25
    .Plot.SeriesCollection(9).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(9).Pen.Width = 25
    .Plot.SeriesCollection(11).Position.Hidden = True
    .Plot.SeriesCollection(11).Pen.Width = 25
    
    .Column = 7
    .ColumnLabel = "2nd Torque(Nm)"
    .Column = 9
    .ColumnLabel = "2nd Power(kW)"
    .Column = 11
    .ColumnLabel = "2nd RPM (1/min x 1000)"
    End If
    
    .Plot.SeriesCollection(5).Position.Hidden = True
    .Plot.SeriesCollection(5).Pen.Width = 25
    
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = EngGraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = EngGraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = EngGraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = RPMXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = RPMXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = RPMXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "1/min x 1000"
    
    Call EngineRPMSmooth(Smoothing + 1)
        
    For i = 1 To LossStart
        
    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    .Row = i
    .Column = 1
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
        .Data = DrumRPM(i) * Transmission / 1000
        Else
        .Data = ClutchRPM / 1000
        End If
    End If
    .Column = 2
    .Data = EngineTorque
    .Column = 3
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
        .Data = DrumRPM(i) * Transmission / 1000
        Else
        .Data = ClutchRPM / 1000
        End If
    End If
    .Column = 4
    .Data = EnginePower
    
    Next i
    
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    For i = 1 To LossStart2
    .Row = i
    .Column = 7
     If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM2nd(i) / 1000
     Else
        .Data = DrumRPM2nd(i) * Transmission2nd / 100
     End If
    .Column = 8
    .Data = EngineTorque2nd(i)
    .Column = 9
     If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM2nd(i) / 1000
     Else
        .Data = DrumRPM2nd(i) * Transmission2nd / 100
     End If
    .Column = 10
    .Data = EnginePower2nd(i)
    Next i
    End If
    
    
        
End Select

End With

End Sub

Public Sub CompareChart()

Erase Speed2nd
Erase Torque2nd
Erase Power2nd
Erase EnginePower2nd
Erase EngineTorque2nd
Erase EngineRPM2nd
Erase DrumRPM2nd
Transmission2nd = 0

With DisplayResults.MSChart1
    
    .Plot.SeriesCollection(7).Pen.VtColor.Set 0, 255, 0
    .Plot.SeriesCollection(9).Pen.VtColor.Set 255, 165, 0
    .Plot.SeriesCollection(11).Pen.VtColor.Set 255, 255, 0


FileDialog:
        CommonDialog1.CancelError = True
        CommonDialog1.Filter = "Open Dyno File (TXT)|*.txt|"
        CommonDialog1.FilterIndex = 1
        
        On Error Resume Next
            CommonDialog1.ShowOpen
        If Err Then Exit Sub
        
        If CommonDialog1.Filename <> "" Then
            On Error Resume Next
            Open CommonDialog1.Filename For Input As #1
            If Err <> 0 Then
                MsgBox "File Open Error", 16, "Error"
                GoTo FileDialog
            End If
            
        
        If CommonDialog1.FilterIndex = 1 Then
                Line Input #1, temp 'version
                Inputstr = Split(temp, Chr(9), 3)
                If Inputstr(0) <> "Version" Then
                    Close #1
                    MsgBox "Invalid File", 16, "Error"
                    Exit Sub
                End If
                DisplayResults.CompareSelection.List(1) = Mid(CommonDialog1.Filename, InStrRev(CommonDialog1.Filename, "\") + 1, Len(CommonDialog1.Filename))
                Line Input #1, temp 'datetime
                Inputstr = Split(temp, Chr(9), 3)
                DateTime2nd = Inputstr(1)
                Line Input #1, temp 'Inertia
                Line Input #1, temp 'Diameter
                Line Input #1, temp 'records
                Inputstr = Split(temp, Chr(9), 3)
                Record2nd = CInt(Inputstr(1)) - Header
                Line Input #1, temp 'Losstart
                Inputstr = Split(temp, Chr(9), 3)
                LossStart2 = CInt(Inputstr(1)) - Header
                Line Input #1, temp 'Loss
                Line Input #1, temp 'RPMPeak
                Line Input #1, temp 'TorquePeak
                Line Input #1, temp 'PowerPeak
                Line Input #1, temp 'Transmission
                Inputstr = Split(temp, Chr(9), 3)
                Transmission2nd = CDbl(Inputstr(1))
                Line Input #1, temp 'Temperature
                Line Input #1, temp 'Pressure
                Line Input #1, temp 'Humidity
                Line Input #1, temp 'Correction
                Line Input #1, temp 'CorrectionType
                Line Input #1, temp 'Description
                Inputstr = Split(temp, Chr(9), 3)
                Description2nd = Inputstr(1)
                Line Input #1, temp 'Displacement
                Inputstr = Split(temp, Chr(9), 3)
                Displacement2nd = Inputstr(1)
                Line Input #1, temp 'Comment
                Inputstr = Split(temp, Chr(9), 3)
                Comment2nd = Inputstr(1)
                Line Input #1, temp
                Line Input #1, temp 'DescriptionRow
            
            i = 0
            
            Do While Not EOF(1)
                    
                    i = i + 1
                    Line Input #1, temp
                    Inputstr = Split(temp, Chr(9), 12)
                    EngineRPM2nd(i) = CDbl(Inputstr(1))
                    If EngineRPM2nd(i) >= MaxEngineRPM2 Then MaxEngineRPM2 = EngineRPM2nd(i)
                    DrumRPM2nd(i) = CDbl(Inputstr(3))
                    If DrumRPM2nd(i) >= MaxRPM2 Then MaxRPM2 = DrumRPM2nd(i)
                    Speed2nd(i) = CDbl(Inputstr(4))
                    If Speed2nd(i) >= MaxSpeed2 Then MaxSpeed2 = Speed2nd(i)
                    
                    Torque2nd(i) = CDbl(Inputstr(5))
                    If i >= Skip And Torque2nd(i) >= MaxTorque2 Then
                        MaxTorque2 = Torque2nd(i)
                        MaxTorqueSpeed2 = Speed2nd(i)
                    End If
                    
                    Power2nd(i) = CDbl(Inputstr(6))
                    If Power2nd(i) >= MaxPower2 Then
                        MaxPower2 = Power2nd(i)
                        MaxPowerSpeed2 = Speed2nd(i)
                    End If
                    
                    EngineTorque2nd(i) = CDbl(Inputstr(9))
                    If EngineTorque2nd(i) >= MaxEngineTorque2 Then
                        MaxEngineTorque2 = EngineTorque2nd(i)
                        If EngineRPM2nd(i) = 0 Then
                            MaxEngineTorqueRPM2 = DrumRPM2nd(i) * Transmission2nd
                        Else
                            MaxEngineTorqueRPM2 = EngineRPM2nd(i)
                        End If
                    End If
                    
                    EnginePower2nd(i) = CDbl(Inputstr(10))
                    If EnginePower2nd(i) >= MaxEnginePower2 Then
                        MaxEnginePower2 = EnginePower2nd(i)
                        If EngineRPM2nd(i) = 0 Then
                            MaxEnginePowerRPM2 = DrumRPM2nd(i) * Transmission2nd
                        Else
                            MaxEnginePowerRPM2 = EngineRPM2nd(i)
                        End If
                    End If
                    
            Loop
    End If
        Close #1
        
Call EngineRPM2Smooth(Smoothing + 1)

If i = Record2nd Then

Call DrawChart

        Else
            Status.Text = "Error"
        End If
    End If
    
End With


End Sub

Public Sub ResetChart()

Unload ComparePrinterSheet

Erase Speed2nd
Erase Torque2nd
Erase Power2nd
Erase EnginePower2nd
Erase EngineTorque2nd
Erase EngineRPM2nd
Erase DrumRPM2nd
Transmission2nd = 0

DisplayResults.CompareSelection.List(1) = "Select File"

With DisplayResults.MSChart1
Select Case DisplayResults.DataSelection.ListIndex
Case 0
    '.ColumnCount = 4
    .Plot.SeriesCollection(5).Position.Hidden = True
    .Plot.SeriesCollection(7).Position.Hidden = True
    .Plot.SeriesCollection(9).Position.Hidden = True
    .Plot.SeriesCollection(11).Position.Hidden = True
Case 1
    '.ColumnCount = 6
    .Plot.SeriesCollection(7).Position.Hidden = True
    .Plot.SeriesCollection(9).Position.Hidden = True
    .Plot.SeriesCollection(11).Position.Hidden = True
Case 2
    .Plot.SeriesCollection(5).Position.Hidden = True
    .Plot.SeriesCollection(7).Position.Hidden = True
    .Plot.SeriesCollection(9).Position.Hidden = True
    .Plot.SeriesCollection(11).Position.Hidden = True
End Select

End With
End Sub


Public Sub PrintChart()

Dim rv As Long

If DisplayResults.CompareSelection.ListIndex <> 1 Then

With PrinterSheet

.WSpeed.Caption = Format(Pi * (Diameter / 1000) * (MaxRPM * 60), "0.00")
.AccelTime.Caption = Format(AccelTime, "0.00")
.WTorque.Caption = Format(MaxTorque, "0.00")
.WTorqueSpeed.Caption = Format(Pi * (Diameter / 1000) * (DrumRPM(TorquePeak) * 60), "0.00")
.WPower.Caption = Format(MaxPower, "0.00") & " (" & Format(MaxPower * 1.359622, "0.00") & "PS)"
.WPowerSpeed.Caption = Format(Pi * (Diameter / 1000) * (DrumRPM(PowerPeak) * 60), "0.00")

If MaxEngineRPM = 0 Then
.ERPM.Caption = Format(MaxRPM * Transmission, "0")
Else
.ERPM.Caption = Format(MaxEngineRPM, "0")
End If

.ETorque.Caption = Format(MaxEngineTorque, "0.00")
If EngineRPM(EngineTorquePeak) = 0 Then
.ETorqueRPM.Caption = Format((DrumRPM(EngineTorquePeak) * Transmission), "0")
.Note.Visible = True
Else
.ETorqueRPM.Caption = Format((EngineRPM(EngineTorquePeak)), "0")
.Note.Visible = False
End If

.EPower.Caption = Format(MaxEnginePower, "0.00") & " (" & Format(MaxEnginePower * 1.359622, "0.00") & "PS)"
If EngineRPM(EnginePowerPeak) = 0 Then
.EPowerRPM.Caption = Format((DrumRPM(EnginePowerPeak) * Transmission), "0")
.Note.Visible = True
Else
.EPowerRPM.Caption = Format(EngineRPM(EnginePowerPeak), "0")
.Note.Visible = False
End If

If StrokeValue.ListIndex = 0 Then .EType.Caption = "2-Stroke" Else .EType.Caption = "4-Stroke"
.Transmission.Caption = Format(Transmission, "0.0000")
.Inertia.Caption = Format(Inertia, "0.0000")
.CalibrationInertia.Caption = Format(CalibrationInertia, "0.0000")
.Diameter.Caption = Format(Diameter, "0.000")
.CorrectionType.Caption = CorrectionType
.CorrectionFactor.Caption = Format(Correction, "0.0000")

.DateTime.Caption = DateTime
.Filename.Caption = Filename.Caption
.Description.Caption = Description
.Displacement.Caption = Displacement
.Comments.Caption = Comment

End With


With ChartPrinterSheet.MSChart1

    chartType = VtChChartType2dXY

    .Plot.AutoLayout = False
    .Plot.UniformAxis = False
    .Plot.WidthToHeightRatio = 1
    .Plot.SeriesCollection(1).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(1).Pen.Width = 25
    .Plot.SeriesCollection(3).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(3).Pen.Width = 25
    
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
    .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    
    
    'title area

    .Title.VtFont.Style = VtFontStyleBold
    .Title.VtFont.Name = "Arial"
    .Title.VtFont.Size = 14
    
    
    'footnote area

    '.FootnoteText = "This is a foot note"
    '.Footnote.VtFont.VtColor.Set 25, 150, 200
    
    '.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"
    '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = "Torque(Nm)"
    '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY2).AxisTitle.Text = "Power(kW)"
    
    .ShowLegend = True

    '.Column = 1
    '.ColumnLabel = "Torque(NM)"
    '.Column = 3
    '.ColumnLabel = "Power(kW)"



Select Case DisplayResults.DataSelection.ListIndex

Case 0

    .Title.Text = "Wheel Data"
    .RowCount = Record
    .ColumnCount = 4
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = GraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = GraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = GraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = SpeedXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = SpeedXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = SpeedXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"

    For i = 1 To Record

    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    .Row = i
    .Column = 1
    .Data = Speed
    .Column = 2
    .Data = Torque
    .Column = 3
    .Data = Speed
    .Column = 4
    .Data = Power
    
    Next i
Case 1

    .Title.Text = "Engine Data"
    .RowCount = LossStart
    .ColumnCount = 6

    .Plot.SeriesCollection(5).Position.Hidden = DisplayResults.Check_RPM.Value
    .Plot.SeriesCollection(5).Pen.Width = 25
    
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    .Column = 5
    .ColumnLabel = "RPM (1/min x 1000)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = EngGraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = EngGraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = EngGraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = SpeedXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = SpeedXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = SpeedXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"
    
    'Call EngineRPMSmooth(Smoothing + 1)
    'Call EngineRPM2Smooth(Smoothing + 1)
    
    For i = 1 To LossStart

    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    .Row = i
    .Column = 1
    .Data = Speed
    .Column = 2
    .Data = EngineTorque
    .Column = 3
    .Data = Speed
    .Column = 4
    .Data = EnginePower
    .Column = 5
    .Data = Speed
    .Column = 6
    
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
            .Data = DrumRPM(i) * Transmission / 1000
        Else
            .Data = ClutchRPM / 1000
        End If
    End If
    
    Next i

Case 2
    .RowCount = LossStart
    
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    .Plot.SeriesCollection(7).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(7).Pen.Width = 25
    .Plot.SeriesCollection(9).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(9).Pen.Width = 25
    .Plot.SeriesCollection(11).Position.Hidden = True
    .Plot.SeriesCollection(11).Pen.Width = 25
    
    .Column = 7
    .ColumnLabel = "2nd Torque(Nm)"
    .Column = 9
    .ColumnLabel = "2nd Power(kW)"
    .Column = 11
    .ColumnLabel = "2nd RPM (1/min x 1000)"
    End If
    
    '.Plot.SeriesCollection(5).Position.Hidden = True
    '.Plot.SeriesCollection(5).Pen.Width = 25
    
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = EngGraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = EngGraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = EngGraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = RPMXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = RPMXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = RPMXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "1/min x 1000"
    
    Call EngineRPMSmooth(Smoothing + 1)
    
    For i = 1 To LossStart

    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)
    
    .Row = i
    .Column = 1
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
        .Data = DrumRPM(i) * Transmission / 1000
        Else
        .Data = ClutchRPM / 1000
        End If
    End If
    .Column = 2
    .Data = EngineTorque
    .Column = 3
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
        .Data = DrumRPM(i) * Transmission / 1000
        Else
        .Data = ClutchRPM / 1000
        End If
    End If
    .Column = 4
    .Data = EnginePower
    Next i
End Select

End With


CommonDialog1.PrinterDefault = True
CommonDialog1.CancelError = True
' Enables error handling to catch cancel error
On Error Resume Next
' display the print dialog box

If Remote = False Then CommonDialog1.ShowPrinter
Remote = False
If Err Then
    ' This code runs if the dialog was cancelled
    'MsgBox "Dialog Cancelled"
    Exit Sub
End If
' Prints the contents of RichTextBox

'DoEvents

'ChartPrinterSheet.Show

'Printer.Orientation = vbPRORLandscape

 
' Make sure picturebox is same size as the chart.
With Picture1
  .Height = ChartPrinterSheet.MSChart1.Height
  .Width = ChartPrinterSheet.MSChart1.Width
End With

Picture1.AutoRedraw = True
rv = SendMessage(ChartPrinterSheet.MSChart1.hwnd, WM_PAINT, Picture1.hDC, 0)
Picture1.Picture = Picture1.Image
Picture1.AutoRedraw = False

' Sent the picture to the clipboard.
Clipboard.Clear
Clipboard.SetData Picture1.Picture

Printer.PaintPicture Picture1.Picture, 0, 6000
PrinterSheet.PrintForm
Printer.EndDoc

Unload PrinterSheet
Unload ChartPrinterSheet

Status.Text = "PRINTED"

Else

' ----------------- Print ComparePrinterSheet


With ComparePrinterSheet


.WSpeed.Caption = Format(Pi * (Diameter / 1000) * (MaxRPM * 60), "0.00")
.AccelTime.Caption = Format(AccelTime, "0.00")
.WTorque.Caption = Format(MaxTorque, "0.00")
.WTorqueSpeed.Caption = Format(Pi * (Diameter / 1000) * (DrumRPM(TorquePeak) * 60), "0.00")
.WPower.Caption = Format(MaxPower, "0.00") & " (" & Format(MaxPower * 1.359622, "0.00") & "PS)"
.WPowerSpeed.Caption = Format(Pi * (Diameter / 1000) * (DrumRPM(PowerPeak) * 60), "0.00")
.WSpeed2.Caption = MaxSpeed2
.WTorque2.Caption = MaxTorque2
.WTorqueSpeed2.Caption = MaxTorqueSpeed2
.WPower2.Caption = Format(MaxPower2, "0.00") & " (" & Format(MaxPower2 * 1.359622, "0.00") & "PS)"
.WPowerSpeed2.Caption = MaxPowerSpeed2
If MaxEngineRPM2 <> 0 Then
    .ERPM2.Caption = Format(MaxEngineRPM2, "0")
Else
    .ERPM2.Caption = Format(MaxRPM2 * Transmission2nd, "0")
End If
.EPower2.Caption = Format(MaxEnginePower2, "0.00") & " (" & Format(MaxEnginePower2 * 1.359622, "0.00") & "PS)"
.EPowerRPM2.Caption = Format(MaxEnginePowerRPM2, "0")
.ETorque2.Caption = MaxEngineTorque2
.ETorqueRPM2.Caption = Format(MaxEngineTorqueRPM2, "0")


If MaxEngineRPM = 0 Then
.ERPM.Caption = Format(MaxRPM * Transmission, "0")
Else
.ERPM.Caption = Format(MaxEngineRPM, "0")
End If

.ETorque.Caption = Format(MaxEngineTorque, "0.00")
If EngineRPM(EngineTorquePeak) = 0 Then
.ETorqueRPM.Caption = Format((DrumRPM(EngineTorquePeak) * Transmission), "0")
.Note.Visible = True
Else
.ETorqueRPM.Caption = Format((EngineRPM(EngineTorquePeak)), "0")
.Note.Visible = False
End If

.EPower.Caption = Format(MaxEnginePower, "0.00") & " (" & Format(MaxEnginePower * 1.359622, "0.00") & "PS)"
If EngineRPM(EnginePowerPeak) = 0 Then
.EPowerRPM.Caption = Format((DrumRPM(EnginePowerPeak) * Transmission), "0")
.Note.Visible = True
Else
.EPowerRPM.Caption = Format(EngineRPM(EnginePowerPeak), "0")
.Note.Visible = False
End If

.DateTime.Caption = DateTime
.Filename.Caption = Filename.Caption
.Description.Caption = Description
.Displacement.Caption = Displacement
.Comments.Caption = Comment
.DateTime2nd.Caption = DateTime2nd
.Filename2nd.Caption = DisplayResults.CompareSelection.List(1)
.Description2nd.Caption = Description2nd
.Displacement2nd.Caption = Displacement2nd
.Comments2nd.Caption = Comment2nd


End With


With ChartPrinterSheet.MSChart1

    chartType = VtChChartType2dXY
    
    .RowCount = 0
    .ColumnCount = 0
    .RandomFill = False

    .ColumnCount = 12
    .Plot.AutoLayout = False
    .Plot.UniformAxis = False
    .Plot.WidthToHeightRatio = 1
    .Plot.SeriesCollection(7).Pen.VtColor.Set 0, 255, 0
    .Plot.SeriesCollection(9).Pen.VtColor.Set 255, 165, 0
    .Plot.SeriesCollection(11).Pen.VtColor.Set 255, 255, 0
    .Plot.SeriesCollection(1).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(1).Pen.Width = 25
    .Plot.SeriesCollection(3).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(3).Pen.Width = 25
    .Plot.SeriesCollection(5).Position.Hidden = True
    .Plot.SeriesCollection(5).Pen.Width = 25
    .Plot.SeriesCollection(7).Position.Hidden = True
    .Plot.SeriesCollection(7).Pen.Width = 25
    .Plot.SeriesCollection(9).Position.Hidden = True
    .Plot.SeriesCollection(9).Pen.Width = 25
    .Plot.SeriesCollection(11).Position.Hidden = True
    .Plot.SeriesCollection(11).Pen.Width = 25
    
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Auto = False
    .Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    
    
    'title area

    .Title.VtFont.Style = VtFontStyleBold
    .Title.VtFont.Name = "Arial"
    .Title.VtFont.Size = 14
    
    
    'footnote area

    '.FootnoteText = "This is a foot note"
    '.Footnote.VtFont.VtColor.Set 25, 150, 200
    
    '.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"
    '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY).AxisTitle.Text = "Torque(Nm)"
    '.Plot.Axis(MSChart20Lib.VtChAxisId.VtChAxisIdY2).AxisTitle.Text = "Power(kW)"
    
    .ShowLegend = True

    '.Column = 1
    '.ColumnLabel = "Torque(Nm)"
    '.Column = 3
    '.ColumnLabel = "Power(kW)"



Select Case DisplayResults.DataSelection.ListIndex

Case 0
    
    ComparePrinterSheet.Engine.Visible = False
    ComparePrinterSheet.Wheel.Visible = True
    ComparePrinterSheet.Engine2.Visible = False
    ComparePrinterSheet.Wheel2.Visible = True
        
    .Plot.SeriesCollection(7).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(9).Position.Hidden = DisplayResults.Check_Power.Value

    .Title.Text = "Wheel Data"
    .RowCount = Record
    If Record2nd > Record Then .RowCount = Record2nd
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    .Column = 7
    .ColumnLabel = "2nd Torque(Nm)"
    .Column = 9
    .ColumnLabel = "2nd Power(kW)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = GraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = GraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = GraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = SpeedXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = SpeedXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = SpeedXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"

    For i = 1 To Record

    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    .Row = i
    .Column = 1
    .Data = Speed
    .Column = 2
    .Data = Torque
    .Column = 3
    .Data = Speed
    .Column = 4
    .Data = Power
    Next i
    
    For i = 1 To Record2nd
    .Row = i
    .Column = 7
    .Data = Speed2nd(i)
    .Column = 8
    .Data = Torque2nd(i)
    .Column = 9
    .Data = Speed2nd(i)
    .Column = 10
    .Data = Power2nd(i)
    Next i
Case 1
    
    ComparePrinterSheet.Engine.Visible = True
    ComparePrinterSheet.Wheel.Visible = False
    ComparePrinterSheet.Engine2.Visible = True
    ComparePrinterSheet.Wheel2.Visible = False
    

    .Title.Text = "Engine Data"
    .RowCount = LossStart
    If LossStart2 > LossStart Then .RowCount = LossStart2
    .Plot.SeriesCollection(5).Position.Hidden = DisplayResults.Check_RPM.Value
    .Plot.SeriesCollection(7).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(9).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(11).Position.Hidden = DisplayResults.Check_RPM.Value
    
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    .Column = 5
    .ColumnLabel = "RPM (1/min x 1000)"
    .Column = 7
    .ColumnLabel = "2nd Torque(Nm)"
    .Column = 9
    .ColumnLabel = "2nd Power(kW)"
    .Column = 11
    .ColumnLabel = "2nd RPM (1/min x 1000)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = EngGraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = EngGraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = EngGraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = SpeedXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = SpeedXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = SpeedXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "km/h"
    
    Call EngineRPMSmooth(Smoothing + 1)
    
    For i = 1 To LossStart

    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    .Row = i
    .Column = 1
    .Data = Speed
    .Column = 2
    .Data = EngineTorque
    .Column = 3
    .Data = Speed
    .Column = 4
    .Data = EnginePower
    .Column = 5
    .Data = Speed
    .Column = 6
    
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
            .Data = DrumRPM(i) * Transmission / 1000
        Else
            .Data = ClutchRPM / 1000
        End If
    End If
    
    Next i
    For i = 1 To LossStart2
    .Row = i
    .Column = 7
    .Data = Speed2nd(i)
    .Column = 8
    .Data = EngineTorque2nd(i)
    .Column = 9
    .Data = Speed2nd(i)
    .Column = 10
    .Data = EnginePower2nd(i)
    .Column = 11
    .Data = Speed2nd(i)
    .Column = 12
    If EngineRPMType.ListIndex = 0 Then
     .Data = SmoothEngineRPM2nd(i) / 1000
    Else
        If DrumRPM2nd(i) * Transmission2nd > ClutchRPM Then
            .Data = DrumRPM2nd(i) * Transmission2nd / 1000
        Else
            .Data = ClutchRPM / 1000
        End If
    End If
    Next i
Case 2
    ComparePrinterSheet.Engine.Visible = True
    ComparePrinterSheet.Wheel.Visible = False
    ComparePrinterSheet.Engine2.Visible = True
    ComparePrinterSheet.Wheel2.Visible = False
    
    .Title.Text = "Engine Data"
    .RowCount = LossStart
    If LossStart2 > LossStart Then .RowCount = LossStart2
    If DisplayResults.CompareSelection.ListIndex = 1 Then
    .Plot.SeriesCollection(7).Position.Hidden = DisplayResults.Check_Torque.Value
    .Plot.SeriesCollection(7).Pen.Width = 25
    .Plot.SeriesCollection(9).Position.Hidden = DisplayResults.Check_Power.Value
    .Plot.SeriesCollection(9).Pen.Width = 25
    .Plot.SeriesCollection(11).Position.Hidden = True
    .Plot.SeriesCollection(11).Pen.Width = 25
    
    .Column = 7
    .ColumnLabel = "2nd Torque(Nm)"
    .Column = 9
    .ColumnLabel = "2nd Power(kW)"
    .Column = 11
    .ColumnLabel = "2nd RPM (1/min x 1000)"
    End If
    
    .Plot.SeriesCollection(5).Position.Hidden = True
    .Plot.SeriesCollection(5).Pen.Width = 25
    
    .Column = 1
    .ColumnLabel = "Torque(Nm)"
    .Column = 3
    .ColumnLabel = "Power(kW)"
    
    .Plot.Axis(VtChAxisIdY).ValueScale.Maximum = EngGraphYMax
    .Plot.Axis(VtChAxisIdY).ValueScale.Minimum = EngGraphYMin
    .Plot.Axis(VtChAxisIdY).ValueScale.MajorDivision = EngGraphYDiv
    .Plot.Axis(VtChAxisIdY).ValueScale.MinorDivision = 1
    
    .Plot.Axis(VtChAxisIdX).ValueScale.Maximum = RPMXMax
    .Plot.Axis(VtChAxisIdX).ValueScale.Minimum = RPMXMin
    .Plot.Axis(VtChAxisIdX).ValueScale.MajorDivision = RPMXDiv
    .Plot.Axis(VtChAxisIdX).ValueScale.MinorDivision = 1
    .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "1/min x 1000"
    
    Call EngineRPMSmooth(Smoothing + 1)
    
    For i = 1 To LossStart

    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)
    
    .Row = i
    .Column = 1
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
        .Data = DrumRPM(i) * Transmission / 1000
        Else
        .Data = ClutchRPM / 1000
        End If
    End If
    .Column = 2
    .Data = EngineTorque
    .Column = 3
    If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM(i) / 1000
    Else
        If DrumRPM(i) * Transmission > ClutchRPM Then
        .Data = DrumRPM(i) * Transmission / 1000
        Else
        .Data = ClutchRPM / 1000
        End If
    End If
    .Column = 4
    .Data = EnginePower
    Next i
    
    For i = 1 To LossStart2
    .Row = i
    .Column = 7
         If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM2nd(i) / 1000
     Else
        .Data = DrumRPM2nd(i) * Transmission2nd / 100
     End If
    .Column = 8
    .Data = EngineTorque2nd(i)
    .Column = 9
         If EngineRPMType.ListIndex = 0 Then
        .Data = SmoothEngineRPM2nd(i) / 1000
     Else
        .Data = DrumRPM2nd(i) * Transmission2nd / 100
     End If
    .Column = 10
    .Data = EnginePower2nd(i)
    Next i
    
End Select



End With

'ComparePrinterSheet.Show

CommonDialog1.PrinterDefault = True
CommonDialog1.CancelError = True
' Enables error handling to catch cancel error
On Error Resume Next
' display the print dialog box

If Remote = False Then CommonDialog1.ShowPrinter
Remote = False
If Err Then
    ' This code runs if the dialog was cancelled
    'MsgBox "Dialog Cancelled"
    Exit Sub
End If

'DoEvents
   
' Make sure picturebox is same size as the chart.
With Picture1
  .Height = ChartPrinterSheet.MSChart1.Height
  .Width = ChartPrinterSheet.MSChart1.Width
End With

Picture1.AutoRedraw = True
rv = SendMessage(ChartPrinterSheet.MSChart1.hwnd, WM_PAINT, Picture1.hDC, 0)
Picture1.Picture = Picture1.Image
Picture1.AutoRedraw = False

' Sent the picture to the clipboard.
Clipboard.Clear
Clipboard.SetData Picture1.Picture

Printer.PaintPicture Picture1.Picture, 0, 6000
ComparePrinterSheet.PrintForm
Printer.EndDoc

Unload ComparePrinterSheet
Unload ChartPrinterSheet

Status.Text = "PRINTED"

End If


End Sub

Public Sub SaveData()

        CommonDialog1.Flags = CommonDialog1.Flags Or FileOpenConstants.cdlOFNOverwritePrompt
        CommonDialog1.CancelError = True
        If Remote = False Then
FileDialog:
            On Error Resume Next
            CommonDialog1.ShowSave
            If Err Then Exit Sub
        End If
        
        If CommonDialog1.Filename <> "" Then
            If DebugLog = True Then
                Print #2, CommonDialog1.Filename
            End If
            On Error Resume Next
            Open CommonDialog1.Filename For Output As #1
            If Err <> 0 Then
                MsgBox "File Open Error", 16, "Error"
                GoTo FileDialog
            End If
                Filename.Caption = Mid(CommonDialog1.Filename, InStrRev(CommonDialog1.Filename, "\") + 1, Len(CommonDialog1.Filename))
                
                Print #1, "Version" & Chr(9); Main.Caption & Chr(9)
                Print #1, "Date / Time" & Chr(9); DateTime & Chr(9)
                Print #1, "Inertia" & Chr(9); Inertia & Chr(9); "CalibrationInertia" & Chr(9); CalibrationInertia & Chr(9); "StaticCalibration" & Chr(9); StaticCalibration & Chr(9); "DynamicCalibration" & Chr(9); DynamicCalibration & Chr(9)
                Print #1, "Diameter" & Chr(9); Diameter & Chr(9)
                Print #1, "LastRecord" & Chr(9); Record + Header & Chr(9)
                Print #1, "LossStart" & Chr(9); LossStart + 1 + Header & Chr(9)
                Select Case LossCalculation.ListIndex
                    Case 0
                    Print #1, "Loss" & Chr(9); "0" & Chr(9)
                    Case 1
                    Print #1, "Loss" & Chr(9); LossPercentage & Chr(9)
                    Case 2
                    Print #1, "Loss" & Chr(9); "None" & Chr(9)
                End Select
                Print #1, "RPMPeak" & Chr(9); RPMPeak + Header & Chr(9)
                Print #1, "TorquePeak" & Chr(9); TorquePeak + Header & Chr(9)
                Print #1, "PowerPeak" & Chr(9); PowerPeak + Header & Chr(9)
                Print #1, "Transmission" & Chr(9); Transmission & Chr(9)
                Print #1, "Temperature" & Chr(9); Temperature & Chr(9)
                Print #1, "Pressure" & Chr(9); Pressure & Chr(9)
                Print #1, "Humidity" & Chr(9); Humidity & Chr(9)
                Print #1, "Correction Factor" & Chr(9); Correction & Chr(9)
                Print #1, "Correction Type" & Chr(9); CorrectionType & Chr(9)
                Print #1, "Name" & Chr(9); Description.Text & Chr(9)
                Print #1, "Displacement" & Chr(9); Displacement.Text & Chr(9)
                Print #1, "Comment" & Chr(9); Comment.Text & Chr(9)
                Print #1,
                
                Print #1, "Time" & Chr(9);
                Print #1, "Engine RPM" & Chr(9);
                Print #1, "Lambda" & Chr(9);
                Print #1, "Drum RPM" & Chr(9);
                Print #1, "Speed(km/h)" & Chr(9);
                Print #1, "Torque(Nm)" & Chr(9);
                Print #1, "Power(kW)" & Chr(9);
                Print #1, "Torque Loss(Nm)" & Chr(9);
                Print #1, "Power Loss(kW)" & Chr(9);
                Print #1, "Engine Torque(Nm)" & Chr(9);
                Print #1, "Engine Power(kW)" & Chr(9)
                
                For i = 1 To Record
                        If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
                        DataCalc (i)
                        Print #1, DrumPeriod & Chr(9);
                        Print #1, Round(EngineRPM(i), 2) & Chr(9);
                        Print #1, Lambda(i) & Chr(9);
                        Print #1, Round(DrumRPM(i), 2) & Chr(9);
                        Print #1, Round(Speed, 2) & Chr(9);
                        Print #1, Round(Torque, 2) & Chr(9);
                        Print #1, Round(Power, 2) & Chr(9);
                        If i <= LossStart Then
                            Print #1, Round(LossTorque, 2) & Chr(9);
                            Print #1, Round(LossPower, 2) & Chr(9);
                            Print #1, Round(EngineTorque, 2) & Chr(9);
                            Print #1, Round(EnginePower, 2);
                        End If
                        Print #1, Chr(9)
                Next i
        Close #1
        Run = Run + 1
        WriteINI "General", "RunNumber", CStr(Run), App.Path & DatabaseExt
        Status.Text = "SAVED"
        End If


End Sub



Private Sub Form_load()


'Set the Smoothing Coefficients for Savitzky-Golay
'The zeroth value is the normalization factor

Main.Caption = "Open Dyno V" & App.Major & "." & App.Minor & App.Revision

DebugLog = False

LoadSetup

EngineRPMType.ListIndex = 0

If START = "NoComm" Then
MsgBox "Com Port Error", 16, "Error"
Unload Me
Else
    If DebugLog = True Then
        Open App.Path & "\OpenDyno.log" For Append As #2
        Status.Text = "DEBUG"
    End If
    SwitchMode ("sensors")
End If

Remote = False

End Sub



Private Sub SaveSetup_Click()

Unload ClimateCorrection
Unload DisplayResults
Unload TransmissionCalc

Stroke = (StrokeValue.ListIndex + 1) * 2
Transmission = CDbl(TransmissionValue.Caption)
Smoothing = CInt(SmoothingValue.Text)



WriteINI "Vehicle", "Stroke", CStr(Stroke), App.Path & DatabaseExt
WriteINI "Vehicle", "Transmission", CStr(Transmission), App.Path & DatabaseExt

WriteINI "Data", "Smoothing", CStr(Smoothing), App.Path & DatabaseExt
WriteINI "Data", "LossMethod", CStr(LossCalculation.ListIndex), App.Path & DatabaseExt
WriteINI "Data", "Correction", CStr(CorrectionFactor.Caption), App.Path & DatabaseExt
WriteINI "Data", "CorrectionType", CStr(CorrectionType), App.Path & DatabaseExt

WriteINI "Data", "EngGraphYMax", CStr(EngGraphYMax), App.Path & DatabaseExt
WriteINI "Data", "EngGraphYMin", CStr(EngGraphYMin), App.Path & DatabaseExt
WriteINI "Data", "EngGraphYDiv", CStr(EngGraphYDiv), App.Path & DatabaseExt
WriteINI "Data", "RPMXMax", CStr(RPMXMax * 1000), App.Path & DatabaseExt
WriteINI "Data", "RPMXMin", CStr(RPMXMin * 1000), App.Path & DatabaseExt
WriteINI "Data", "RPMXDiv", CStr(RPMXDiv), App.Path & DatabaseExt
WriteINI "Data", "GraphYMax", CStr(GraphYMax), App.Path & DatabaseExt
WriteINI "Data", "GraphYMin", CStr(GraphYMin), App.Path & DatabaseExt
WriteINI "Data", "GraphYDiv", CStr(GraphYDiv), App.Path & DatabaseExt
WriteINI "Data", "SpeedXMax", CStr(SpeedXMax), App.Path & DatabaseExt
WriteINI "Data", "SpeedXMin", CStr(SpeedXMin), App.Path & DatabaseExt
WriteINI "Data", "SpeedXDiv", CStr(SpeedXDiv), App.Path & DatabaseExt


End Sub


Private Sub LoadSetup()

Dim temp As String

CommonDialog1.InitDir = ReadINI("General", "DataPath", App.Path & DatabaseExt)
On Error Resume Next
DebugLog = CBool(ReadINI("General", "DebugLog", App.Path & DatabaseExt))
AnalyzeWith = ReadINI("General", "AnalyzeWith", App.Path & DatabaseExt)
AppendFilename = CBool(ReadINI("General", "AppendFilename", App.Path & DatabaseExt))
Run = CInt(ReadINI("General", "RunNumber", App.Path & DatabaseExt))

Pulse = CLng(ReadINI("Dyno", "Pulse", App.Path & DatabaseExt))
PulseValue.Caption = Pulse
Period = CDbl(ReadINI("Dyno", "Period", App.Path & DatabaseExt))
Inertia = CDbl(ReadINI("Dyno", "Inertia", App.Path & DatabaseExt))
InertiaValue.Caption = Inertia
CalibrationInertia = CDbl(ReadINI("Dyno", "CalibrationInertia", App.Path & DatabaseExt))
CalibrationInertiaValue.Caption = CalibrationInertia
StaticCalibration = CDbl(ReadINI("Dyno", "StaticCalibration", App.Path & DatabaseExt))
DynamicCalibration = CDbl(ReadINI("Dyno", "DynamicCalibration", App.Path & DatabaseExt))
Diameter = CDbl(ReadINI("Dyno", "Diameter", App.Path & DatabaseExt))
DiameterValue.Caption = Diameter
YellowRPM = CInt(ReadINI("Dyno", "YellowRPM", App.Path & DatabaseExt))
RedRPM = CInt(ReadINI("Dyno", "RedRPM", App.Path & DatabaseExt))

Transmission = CDbl(ReadINI("Vehicle", "Transmission", App.Path & DatabaseExt))
TransmissionValue.Caption = Round(Transmission, 2)
Stroke = CInt(ReadINI("Vehicle", "Stroke", App.Path & DatabaseExt))
StrokeValue.ListIndex = Round((Stroke / 2) - 1)
'If CBool(ReadINI("Vehicle", "EngineRPM", App.Path & DatabaseExt)) = True Then
'EngineRPMType.ListIndex = 0
'Clutch.Text = 0
'ClutchRPM = 0
'Clutch.Enabled = False
'TransmissionValue.Enabled = False
'Else
'EngineRPMType.ListIndex = 1
'End If


Smoothing = CInt(ReadINI("Data", "Smoothing", App.Path & DatabaseExt))
SmoothingValue.Text = Smoothing
GraphYMax = CInt(ReadINI("Data", "GraphYMax", App.Path & DatabaseExt))
GraphYMin = CInt(ReadINI("Data", "GraphYMin", App.Path & DatabaseExt))
GraphYDiv = CInt(ReadINI("Data", "GraphYDiv", App.Path & DatabaseExt))
EngGraphYMax = CInt(ReadINI("Data", "EngGraphYMax", App.Path & DatabaseExt))
EngGraphYMin = CInt(ReadINI("Data", "EngGraphYMin", App.Path & DatabaseExt))
EngGraphYDiv = CInt(ReadINI("Data", "EngGraphYDiv", App.Path & DatabaseExt))
SpeedXMax = CInt(ReadINI("Data", "SpeedXMax", App.Path & DatabaseExt))
SpeedXMin = CInt(ReadINI("Data", "SpeedXMin", App.Path & DatabaseExt))
SpeedXDiv = CInt(ReadINI("Data", "SpeedXDiv", App.Path & DatabaseExt))
RPMXMax = CInt(ReadINI("Data", "RPMXMax", App.Path & DatabaseExt)) / 1000
RPMXMin = CInt(ReadINI("Data", "RPMXMin", App.Path & DatabaseExt)) / 1000
RPMXDiv = CInt(ReadINI("Data", "RPMXDiv", App.Path & DatabaseExt))
LossCalculation.ListIndex = CInt(ReadINI("Data", "LossMethod", App.Path & DatabaseExt))
LossPercentage = CInt(ReadINI("Data", "LossPercentage", App.Path & DatabaseExt))
CorrectionFactor.Caption = Format(ReadINI("Data", "Correction", App.Path & DatabaseExt), "0.0000")
Correction = CDbl(CorrectionFactor.Caption)
CorrectionType = (ReadINI("Data", "CorrectionType", App.Path & DatabaseExt))
LambdaEnable = CBool(ReadINI("Lambda", "LambdaEnable", App.Path & DatabaseExt))
For i = 65 To 121
    LambdaFactor(i) = CDbl(ReadINI("Lambda", "Lambda" + CStr(i), App.Path & DatabaseExt))
Next i
End Sub

Private Sub GetPeak()

MaxRPM = 0
MaxEngineRPM = 0
MaxTorque = 0
MaxPower = 0
LossStart = 0
TorquePeak = 0
PowerPeak = 0
MaxEngineTorque = 0
MaxEnginePower = 0
EngineTorquePeak = 0
EnginePowerPeak = 0
AccelTime = 0

For i = Skip To Record
    
    If DrumRPM(i) <> 0 Then DrumPeriod = 1 / (DrumRPM(i) / 60) Else DrumPeriod = 0
    DataCalc (i)

    If DrumRPM(i) >= MaxRPM Then
        MaxRPM = DrumRPM(i)
        RPMPeak = i
    End If
    
  If i <= LossStart Then
    
    If EngineRPMType.ListIndex = 0 Then
        If EngineRPM(i) >= MaxEngineRPM Then MaxEngineRPM = EngineRPM(i)
    Else
        If DrumRPM(i) * Transmission >= MaxEngineRPM Then MaxEngineRPM = DrumRPM(i) * Transmission
    End If
    
    If Torque >= MaxTorque Then
        TorquePeak = i
        MaxTorque = Torque
    End If

    If Power >= MaxPower Then
        PowerPeak = i
        MaxPower = Power
    End If

    If EngineTorque >= MaxEngineTorque Then
        EngineTorquePeak = i
        MaxEngineTorque = EngineTorque
    End If

    If EnginePower >= MaxEnginePower Then
        EnginePowerPeak = i
        MaxEnginePower = EnginePower
    End If
  
  End If
 
Next i

For i = 1 To RPMPeak
    If DrumRPM(i) <> 0 Then AccelTime = AccelTime + (60 / DrumRPM(i))
Next i

If EngineRPMType.ListIndex = 0 Then
    If EngineRPM(PowerPeak) <> 0 Then
        Transmission = EngineRPM(PowerPeak) / DrumRPM(PowerPeak)
        TransmissionValue.Caption = Round(Transmission, 2)
    End If
End If
End Sub

Private Sub EngineRPMCalc_alternativ() ' Take max raw value

Erase EngineRPM

J = 1
For i = 0 To DataNo
If RawEngine(i) > 200 And RawEngine(i) < 12000 Then
temp = (60 / (RawEngine(i) * Period)) / (Stroke / 2)
Else
temp = 0
End If
If temp > EngineRPM(J) Then EngineRPM(J) = temp
If ((i + 1) Mod Pulse) = 0 Then
    J = J + 1
End If
Next i

End Sub

Private Sub EngineRPMCalc() 'Average Raw values

Erase EngineRPM

J = 1
For i = 0 To DataNo
        If RawEngine(i) <> 0 Then EngineRPM(J) = EngineRPM(J) + ((60 / (RawEngine(i) * Period)) / (Stroke / 2))
        If ((i + 1) Mod Pulse) = 0 Then
            EngineRPM(J) = EngineRPM(J) / Pulse
            J = J + 1
        End If
Next i

End Sub


Private Sub LambdaCalc()

Erase Lambda
If LambdaEnable = True Then

    J = 1
    For i = 0 To DataNo
        Lambda(J) = Lambda(J) + RawLambda(i)
        If ((i + 1) Mod Pulse) = 0 Then
            Lambda(J) = Lambda(J) / Pulse
            k = 65
            Do While Lambda(J) < LambdaFactor(k)
                k = k + 1
            Loop
            Lambda(J) = Format(k / 100, "0.00")
            J = J + 1
        End If
    Next i

End If


End Sub


Private Sub DrumRPMCalc()

Erase DrumRPM

Record = 0
DrumPeriod = 0
                
For i = 0 To DataNo
    DrumPeriod = DrumPeriod + (RawDyno(i) * Period)
    If ((i + 1) Mod Pulse) = 0 And DrumPeriod <> 0 Then
        Record = Record + 1
        DrumRPM(Record) = (1 / DrumPeriod) * 60
        DrumPeriod = 0
    End If
Next i

Record = Record - 1

End Sub

Private Sub AlphaCalc(Level As Integer)

'Savitzky_Golay Smoothing

'The Savitzky-Golay smoothing algorithm essentialy fits the data to a second order polynomial
'within a moving data window.  It assumes that the data has a fixed spacing in the x direction,
'but does work even if this is not the case.

'For more info see:
'"Smoothing and Differentiation of Data by Simplified Least Squares Procedure",
'Abraham Savitzky and Marcel J. E. Golay, Analytical Chemistry, Vol. 36, No. 8, Page 1627 (1964)

'Degree 2 = 5 point
'Degree 3 = 7 point ...etc

Dim i As Integer, J As Integer, k As Integer
Dim TempSum As Double
Dim Degree As Integer
On Error Resume Next


'The matrix for the Savitzky-Golay Coefficents

Dim SGCoef(1 To 11, 0 To 13) As Integer

SGCoef(1, 1) = 17
SGCoef(1, 2) = 12
SGCoef(1, 3) = -3
SGCoef(1, 0) = 35

SGCoef(2, 1) = 7
SGCoef(2, 2) = 6
SGCoef(2, 3) = 3
SGCoef(2, 4) = -2
SGCoef(2, 0) = 21

SGCoef(3, 1) = 59
SGCoef(3, 2) = 54
SGCoef(3, 3) = 39
SGCoef(3, 4) = 14
SGCoef(3, 5) = -21
SGCoef(3, 0) = 231

SGCoef(4, 1) = 89
SGCoef(4, 2) = 84
SGCoef(4, 3) = 69
SGCoef(4, 4) = 44
SGCoef(4, 5) = 9
SGCoef(4, 6) = -36
SGCoef(4, 0) = 429


SGCoef(5, 1) = 25
SGCoef(5, 2) = 24
SGCoef(5, 3) = 21
SGCoef(5, 4) = 16
SGCoef(5, 5) = 9
SGCoef(5, 6) = 0
SGCoef(5, 7) = -11
SGCoef(5, 0) = 143

SGCoef(6, 1) = 167
SGCoef(6, 2) = 162
SGCoef(6, 3) = 147
SGCoef(6, 4) = 122
SGCoef(6, 5) = 87
SGCoef(6, 6) = 42
SGCoef(6, 7) = -13
SGCoef(6, 8) = -78
SGCoef(6, 0) = 1105

SGCoef(7, 1) = 43
SGCoef(7, 2) = 42
SGCoef(7, 3) = 39
SGCoef(7, 4) = 34
SGCoef(7, 5) = 27
SGCoef(7, 6) = 18
SGCoef(7, 7) = 7
SGCoef(7, 8) = -6
SGCoef(7, 9) = -21
SGCoef(7, 0) = 323

SGCoef(8, 1) = 269
SGCoef(8, 2) = 264
SGCoef(8, 3) = 249
SGCoef(8, 4) = 224
SGCoef(8, 5) = 189
SGCoef(8, 6) = 144
SGCoef(8, 7) = 89
SGCoef(8, 8) = 24
SGCoef(8, 9) = -51
SGCoef(8, 10) = -136
SGCoef(8, 0) = 2261

SGCoef(9, 1) = 329
SGCoef(9, 2) = 324
SGCoef(9, 3) = 309
SGCoef(9, 4) = 284
SGCoef(9, 5) = 249
SGCoef(9, 6) = 204
SGCoef(9, 7) = 149
SGCoef(9, 8) = 84
SGCoef(9, 9) = 9
SGCoef(9, 10) = -76
SGCoef(9, 11) = -171
SGCoef(9, 0) = 3059

SGCoef(10, 1) = 79
SGCoef(10, 2) = 78
SGCoef(10, 3) = 75
SGCoef(10, 4) = 70
SGCoef(10, 5) = 63
SGCoef(10, 6) = 54
SGCoef(10, 7) = 43
SGCoef(10, 8) = 30
SGCoef(10, 9) = 15
SGCoef(10, 10) = -2
SGCoef(10, 11) = -21
SGCoef(10, 12) = -42
SGCoef(10, 0) = 806

SGCoef(11, 1) = 467
SGCoef(11, 2) = 462
SGCoef(11, 3) = 447
SGCoef(11, 4) = 422
SGCoef(11, 5) = 387
SGCoef(11, 6) = 322
SGCoef(11, 7) = 287
SGCoef(11, 8) = 222
SGCoef(11, 9) = 147
SGCoef(11, 10) = 62
SGCoef(11, 11) = -33
SGCoef(11, 12) = -138
SGCoef(11, 13) = -253
SGCoef(11, 0) = 5135

Erase Alpha
Erase LossAlpha


For i = 1 To Record
If DrumRPM(i - 1) <> 0 And DrumRPM(i) <> 0 Then Alpha(i) = 2 * Pi * ((DrumRPM(i) / 60) - (DrumRPM(i - 1) / 60)) * (DrumRPM(i) / 60) Else Alpha(i) = 0
Next i

StartSmoothing:

If Level > 12 Then
    Degree = 12
Else
    Degree = Level
End If

Level = Level - Degree


  'we cannot smooth too close to the data bounds
  
  For i = 1 + Degree To Record - Degree
    TempSum = Alpha(i) * SGCoef(Degree - 1, 1)
    For J = 1 To Degree
      TempSum = TempSum + Alpha(i - J) * (SGCoef(Degree - 1, J + 1))
      TempSum = TempSum + Alpha(i + J) * (SGCoef(Degree - 1, J + 1))
    Next J
    Alpha(i) = TempSum / SGCoef(Degree - 1, 0)
  Next i
  
  'The last smoothed data will be used to create a new smoothed data set,
  'therefore the smoothing operations will be additive
  
If Level > 0 Then GoTo StartSmoothing

    For i = 1 To Record
        For J = Record To 1 Step -1
            If (DrumRPM(i) >= DrumRPM(J)) And (Alpha(J) < 0) Then LossAlpha(i) = Alpha(J)
        Next J
    Next i


End Sub

Private Sub EngineRPM2Smooth(Level As Integer)

'Savitzky_Golay Smoothing

'The Savitzky-Golay smoothing algorithm essentialy fits the data to a second order polynomial
'within a moving data window.  It assumes that the data has a fixed spacing in the x direction,
'but does work even if this is not the case.

'For more info see:
'"Smoothing and Differentiation of Data by Simplified Least Squares Procedure",
'Abraham Savitzky and Marcel J. E. Golay, Analytical Chemistry, Vol. 36, No. 8, Page 1627 (1964)

'Degree 2 = 5 point
'Degree 3 = 7 point ...etc

Dim i As Integer, J As Integer, k As Integer
Dim TempSum As Double
Dim Degree As Integer
On Error Resume Next


'The matrix for the Savitzky-Golay Coefficents

Dim SGCoef(1 To 11, 0 To 13) As Integer

SGCoef(1, 1) = 17
SGCoef(1, 2) = 12
SGCoef(1, 3) = -3
SGCoef(1, 0) = 35

SGCoef(2, 1) = 7
SGCoef(2, 2) = 6
SGCoef(2, 3) = 3
SGCoef(2, 4) = -2
SGCoef(2, 0) = 21

SGCoef(3, 1) = 59
SGCoef(3, 2) = 54
SGCoef(3, 3) = 39
SGCoef(3, 4) = 14
SGCoef(3, 5) = -21
SGCoef(3, 0) = 231

SGCoef(4, 1) = 89
SGCoef(4, 2) = 84
SGCoef(4, 3) = 69
SGCoef(4, 4) = 44
SGCoef(4, 5) = 9
SGCoef(4, 6) = -36
SGCoef(4, 0) = 429


SGCoef(5, 1) = 25
SGCoef(5, 2) = 24
SGCoef(5, 3) = 21
SGCoef(5, 4) = 16
SGCoef(5, 5) = 9
SGCoef(5, 6) = 0
SGCoef(5, 7) = -11
SGCoef(5, 0) = 143

SGCoef(6, 1) = 167
SGCoef(6, 2) = 162
SGCoef(6, 3) = 147
SGCoef(6, 4) = 122
SGCoef(6, 5) = 87
SGCoef(6, 6) = 42
SGCoef(6, 7) = -13
SGCoef(6, 8) = -78
SGCoef(6, 0) = 1105

SGCoef(7, 1) = 43
SGCoef(7, 2) = 42
SGCoef(7, 3) = 39
SGCoef(7, 4) = 34
SGCoef(7, 5) = 27
SGCoef(7, 6) = 18
SGCoef(7, 7) = 7
SGCoef(7, 8) = -6
SGCoef(7, 9) = -21
SGCoef(7, 0) = 323

SGCoef(8, 1) = 269
SGCoef(8, 2) = 264
SGCoef(8, 3) = 249
SGCoef(8, 4) = 224
SGCoef(8, 5) = 189
SGCoef(8, 6) = 144
SGCoef(8, 7) = 89
SGCoef(8, 8) = 24
SGCoef(8, 9) = -51
SGCoef(8, 10) = -136
SGCoef(8, 0) = 2261

SGCoef(9, 1) = 329
SGCoef(9, 2) = 324
SGCoef(9, 3) = 309
SGCoef(9, 4) = 284
SGCoef(9, 5) = 249
SGCoef(9, 6) = 204
SGCoef(9, 7) = 149
SGCoef(9, 8) = 84
SGCoef(9, 9) = 9
SGCoef(9, 10) = -76
SGCoef(9, 11) = -171
SGCoef(9, 0) = 3059

SGCoef(10, 1) = 79
SGCoef(10, 2) = 78
SGCoef(10, 3) = 75
SGCoef(10, 4) = 70
SGCoef(10, 5) = 63
SGCoef(10, 6) = 54
SGCoef(10, 7) = 43
SGCoef(10, 8) = 30
SGCoef(10, 9) = 15
SGCoef(10, 10) = -2
SGCoef(10, 11) = -21
SGCoef(10, 12) = -42
SGCoef(10, 0) = 806

SGCoef(11, 1) = 467
SGCoef(11, 2) = 462
SGCoef(11, 3) = 447
SGCoef(11, 4) = 422
SGCoef(11, 5) = 387
SGCoef(11, 6) = 322
SGCoef(11, 7) = 287
SGCoef(11, 8) = 222
SGCoef(11, 9) = 147
SGCoef(11, 10) = 62
SGCoef(11, 11) = -33
SGCoef(11, 12) = -138
SGCoef(11, 13) = -253
SGCoef(11, 0) = 5135


For i = 0 To 9999
SmoothEngineRPM2nd(i) = EngineRPM2nd(i)
Next i

StartSmoothing:

If Level > 12 Then
    Degree = 12
Else
    Degree = Level
End If

Level = Level - Degree


  'we cannot smooth too close to the data bounds
  
  For i = 1 + Degree To Record - Degree
    TempSum = SmoothEngineRPM2nd(i) * SGCoef(Degree - 1, 1)
    For J = 1 To Degree
      TempSum = TempSum + SmoothEngineRPM2nd(i - J) * (SGCoef(Degree - 1, J + 1))
      TempSum = TempSum + SmoothEngineRPM2nd(i + J) * (SGCoef(Degree - 1, J + 1))
    Next J
    SmoothEngineRPM2nd(i) = TempSum / SGCoef(Degree - 1, 0)
  Next i
  
  'The last smoothed data will be used to create a new smoothed data set,
  'therefore the smoothing operations will be additive
  
If Level > 0 Then GoTo StartSmoothing
  
End Sub


Private Sub EngineRPMSmooth(Level As Integer)

'Savitzky_Golay Smoothing

'The Savitzky-Golay smoothing algorithm essentialy fits the data to a second order polynomial
'within a moving data window.  It assumes that the data has a fixed spacing in the x direction,
'but does work even if this is not the case.

'For more info see:
'"Smoothing and Differentiation of Data by Simplified Least Squares Procedure",
'Abraham Savitzky and Marcel J. E. Golay, Analytical Chemistry, Vol. 36, No. 8, Page 1627 (1964)

'Degree 2 = 5 point
'Degree 3 = 7 point ...etc

Dim i As Integer, J As Integer, k As Integer
Dim TempSum As Double
Dim Degree As Integer
On Error Resume Next


'The matrix for the Savitzky-Golay Coefficents

Dim SGCoef(1 To 11, 0 To 13) As Integer

SGCoef(1, 1) = 17
SGCoef(1, 2) = 12
SGCoef(1, 3) = -3
SGCoef(1, 0) = 35

SGCoef(2, 1) = 7
SGCoef(2, 2) = 6
SGCoef(2, 3) = 3
SGCoef(2, 4) = -2
SGCoef(2, 0) = 21

SGCoef(3, 1) = 59
SGCoef(3, 2) = 54
SGCoef(3, 3) = 39
SGCoef(3, 4) = 14
SGCoef(3, 5) = -21
SGCoef(3, 0) = 231

SGCoef(4, 1) = 89
SGCoef(4, 2) = 84
SGCoef(4, 3) = 69
SGCoef(4, 4) = 44
SGCoef(4, 5) = 9
SGCoef(4, 6) = -36
SGCoef(4, 0) = 429


SGCoef(5, 1) = 25
SGCoef(5, 2) = 24
SGCoef(5, 3) = 21
SGCoef(5, 4) = 16
SGCoef(5, 5) = 9
SGCoef(5, 6) = 0
SGCoef(5, 7) = -11
SGCoef(5, 0) = 143

SGCoef(6, 1) = 167
SGCoef(6, 2) = 162
SGCoef(6, 3) = 147
SGCoef(6, 4) = 122
SGCoef(6, 5) = 87
SGCoef(6, 6) = 42
SGCoef(6, 7) = -13
SGCoef(6, 8) = -78
SGCoef(6, 0) = 1105

SGCoef(7, 1) = 43
SGCoef(7, 2) = 42
SGCoef(7, 3) = 39
SGCoef(7, 4) = 34
SGCoef(7, 5) = 27
SGCoef(7, 6) = 18
SGCoef(7, 7) = 7
SGCoef(7, 8) = -6
SGCoef(7, 9) = -21
SGCoef(7, 0) = 323

SGCoef(8, 1) = 269
SGCoef(8, 2) = 264
SGCoef(8, 3) = 249
SGCoef(8, 4) = 224
SGCoef(8, 5) = 189
SGCoef(8, 6) = 144
SGCoef(8, 7) = 89
SGCoef(8, 8) = 24
SGCoef(8, 9) = -51
SGCoef(8, 10) = -136
SGCoef(8, 0) = 2261

SGCoef(9, 1) = 329
SGCoef(9, 2) = 324
SGCoef(9, 3) = 309
SGCoef(9, 4) = 284
SGCoef(9, 5) = 249
SGCoef(9, 6) = 204
SGCoef(9, 7) = 149
SGCoef(9, 8) = 84
SGCoef(9, 9) = 9
SGCoef(9, 10) = -76
SGCoef(9, 11) = -171
SGCoef(9, 0) = 3059

SGCoef(10, 1) = 79
SGCoef(10, 2) = 78
SGCoef(10, 3) = 75
SGCoef(10, 4) = 70
SGCoef(10, 5) = 63
SGCoef(10, 6) = 54
SGCoef(10, 7) = 43
SGCoef(10, 8) = 30
SGCoef(10, 9) = 15
SGCoef(10, 10) = -2
SGCoef(10, 11) = -21
SGCoef(10, 12) = -42
SGCoef(10, 0) = 806

SGCoef(11, 1) = 467
SGCoef(11, 2) = 462
SGCoef(11, 3) = 447
SGCoef(11, 4) = 422
SGCoef(11, 5) = 387
SGCoef(11, 6) = 322
SGCoef(11, 7) = 287
SGCoef(11, 8) = 222
SGCoef(11, 9) = 147
SGCoef(11, 10) = 62
SGCoef(11, 11) = -33
SGCoef(11, 12) = -138
SGCoef(11, 13) = -253
SGCoef(11, 0) = 5135


For i = 0 To 9999
SmoothEngineRPM(i) = EngineRPM(i)
Next i

StartSmoothing:

If Level > 12 Then
    Degree = 12
Else
    Degree = Level
End If

Level = Level - Degree


  'we cannot smooth too close to the data bounds
  
  For i = 1 + Degree To Record - Degree
    TempSum = SmoothEngineRPM(i) * SGCoef(Degree - 1, 1)
    For J = 1 To Degree
      TempSum = TempSum + SmoothEngineRPM(i - J) * (SGCoef(Degree - 1, J + 1))
      TempSum = TempSum + SmoothEngineRPM(i + J) * (SGCoef(Degree - 1, J + 1))
    Next J
    SmoothEngineRPM(i) = TempSum / SGCoef(Degree - 1, 0)
  Next i
  
  'The last smoothed data will be used to create a new smoothed data set,
  'therefore the smoothing operations will be additive
  
If Level > 0 Then GoTo StartSmoothing
  
End Sub


Private Sub TransmissionValue_Click()
TransmissionCalc.MaxEngineRPM.Text = Format(MaxEngineRPM, "0")
TransmissionCalc.TransmissionValue.Text = Transmission
TransmissionCalc.Show
TransmissionCalc.TransmissionValue.Text = TransmissionValue.Caption
End Sub



Public Sub LoadData_Click()

Dim temp As String

If Mode = "record" And DataNo > Pulse Then
    SwitchMode ("save")
Else
    SwitchMode ("standby")


Call LoadSetup

Pulse = 1
PulseValue.Caption = Pulse

FileDialog:
        CommonDialog1.CancelError = True
        CommonDialog1.Filter = "Open Dyno File (TXT)|*.txt|" & "WTK Datei (WTK)|*.wtk"
        CommonDialog1.FilterIndex = 1
        
        On Error Resume Next
            CommonDialog1.ShowOpen
        If Err Then Exit Sub
        
        Unload ClimateCorrection
        Unload DisplayResults
        Unload TransmissionCalc
        
        If CommonDialog1.Filename <> "" Then
            On Error Resume Next
            Open CommonDialog1.Filename For Input As #1
            If Err <> 0 Then
                MsgBox "File Open Error", 16, "Error"
                GoTo FileDialog
            End If
                Filename.Caption = Mid(CommonDialog1.Filename, InStrRev(CommonDialog1.Filename, "\") + 1, Len(CommonDialog1.Filename))
                
        If CommonDialog1.FilterIndex = 1 Then
                Line Input #1, temp 'version
                Inputstr = Split(temp, Chr(9), 3)
                If Inputstr(0) <> "Version" Then
                    Close #1
                    MsgBox "Invalid File", 16, "Error"
                    Exit Sub
                End If
                Line Input #1, temp 'datetime
                Inputstr = Split(temp, Chr(9), 3)
                DateTime = Inputstr(1)
                Line Input #1, temp 'Inertia
                Inputstr = Split(temp, Chr(9), 8)
                Inertia = CDbl(Inputstr(1))
                InertiaValue.Caption = Inertia
                If CDbl(Inputstr(3)) > 0 Then
                   CalibrationInertia = CDbl(Inputstr(3))
                   CalibrationInertiaValue.Caption = CalibrationInertia
                    StaticCalibration = CDbl(Inputstr(5))
                    DynamicCalibration = CDbl(Inputstr(7))
                End If
                Line Input #1, temp 'Diameter
                Inputstr = Split(temp, Chr(9), 3)
                Diameter = CDbl(Inputstr(1))
                DiameterValue.Caption = Diameter
                Line Input #1, temp 'records
                Inputstr = Split(temp, Chr(9), 3)
                Record = CInt(Inputstr(1)) - Header
                Line Input #1, temp 'Losstart
                'Inputstr = Split(temp, Chr(9), 3)
                'LossStart = CInt(Inputstr(1)) - Header
                Line Input #1, temp 'Loss
                Inputstr = Split(temp, Chr(9), 3)
                If Inputstr(1) <> "None" Then
                    If CInt(Inputstr(1)) = 0 Then
                        LossCalculation.ListIndex = 0
                    Else
                        LossCalculation.ListIndex = 1
                        Loss = CInt(Inputstr(1))
                    End If
                Else
                    LossCalculation.ListIndex = 2
                End If
                Line Input #1, temp 'RPMPeak
                'Inputstr = Split(temp, Chr(9), 3)
                'RPMPeak = CInt(Inputstr(1)) - Header
                Line Input #1, temp 'TorquePeak
                'Inputstr = Split(temp, Chr(9), 3)
                'TorquePeak = CInt(Inputstr(1)) - Header
                Line Input #1, temp 'PowerPeak
                'Inputstr = Split(temp, Chr(9), 3)
                'PowerPeak = CInt(Inputstr(1)) - Header
                Line Input #1, temp 'Transmission
                Inputstr = Split(temp, Chr(9), 3)
                Transmission = CDbl(Inputstr(1))
                TransmissionValue.Caption = Round(Transmission, 2)
                Line Input #1, temp 'Temperature
                Inputstr = Split(temp, Chr(9), 3)
                Temperature = CDbl(Inputstr(1))
                Line Input #1, temp 'Pressure
                Inputstr = Split(temp, Chr(9), 3)
                Pressure = CInt(Inputstr(1))
                Line Input #1, temp 'Humidity
                Inputstr = Split(temp, Chr(9), 3)
                Humidity = CDbl(Inputstr(1))
                Line Input #1, temp 'Correction
                Inputstr = Split(temp, Chr(9), 3)
                Correction = CDbl(Inputstr(1))
                CorrectionFactor.Caption = Format(Correction, "0.0000")
                Line Input #1, temp 'CorrectionType
                Inputstr = Split(temp, Chr(9), 3)
                CorrectionType = Inputstr(1)
                Line Input #1, temp 'Description
                Inputstr = Split(temp, Chr(9), 3)
                Description.Text = Inputstr(1)
                Line Input #1, temp 'Displacement
                Inputstr = Split(temp, Chr(9), 3)
                Displacement.Text = Inputstr(1)
                Line Input #1, temp 'Comment
                Inputstr = Split(temp, Chr(9), 3)
                Comment.Text = Inputstr(1)
                Line Input #1, temp
                Line Input #1, temp 'DescriptionRow
            
            i = 0
            
            Do While Not EOF(1)
                    
                    i = i + 1
                    Line Input #1, temp
                    Inputstr = Split(temp, Chr(9), 12)
                    EngineRPM(i) = CDbl(Inputstr(1))
                    Lambda(i) = CDbl(Inputstr(2))
                    If CDbl(Inputstr(0)) <> 0 Then DrumRPM(i) = (1 / CDbl(Inputstr(0))) * 60 Else DrumRPM(i) = 0
            Loop
    Else
            Line Input #1, temp 'check wtk header
            If temp <> "-------------------------------------------------------------------------------" Then
                Close #1
                MsgBox "Invalid File", 16, "Error"
            Exit Sub
            End If
            For i = 1 To 91
            Line Input #1, temp 'skip header
            Next i
            Line Input #1, temp '
            Record = CInt(temp) - 3
            For i = 1 To 4
            Line Input #1, temp 'jump to data
            Next i
            
            i = 0
            
            Do While Not EOF(1)
                    i = i + 1
                    Line Input #1, temp
                    Inputstr = Split(temp, ";", 3)
                    If CLng(Inputstr(0)) <> 0 Then DrumRPM(i) = (1 / Round((CLng(Inputstr(0)) * 0.0000001), 5)) * 60 Else DrumRPM(i) = 0
            Loop
            For k = i To i + 12
            DrumRPM(k) = DrumRPM(Record)
            Next k
            
            i = i + 12
            Record = Record + 12
            
            CommonDialog1.FilterIndex = 1
            Comment.Text = Filename.Caption
            Description.Text = "Converted from WTK"
            Displacement.Text = ""
            
    End If
        Close #1
        If i = Record Then
            Status.Text = "LOADED"
            Call AlphaCalc(Smoothing + 1)
            Call GetPeak
            Call DisplayResult
        Else
            Status.Text = "Error"
        End If
    End If
End If
    

End Sub

Private Sub ShellAndWait(ByVal program_name As String, _
    ByVal window_style As VbAppWinStyle)



Dim process_id As Long
Dim process_handle As Long

    ' Start the program.
    On Error GoTo ShellError
    process_id = Shell(program_name, window_style)
    On Error GoTo 0

    ' Hide.
    Me.Visible = False
    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    ' Reappear.
    Me.Visible = True
    Exit Sub

ShellError:
    MsgBox "Error starting task " & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub

