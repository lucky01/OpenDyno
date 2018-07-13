VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DynoTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Open Dyno V0.1"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   12525
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "csv"
   End
   Begin VB.Frame Frame3 
      Height          =   1575
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "MODE1"
         Height          =   495
         Index           =   2
         Left            =   7200
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MODE0"
         Height          =   495
         Index           =   2
         Left            =   5520
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Lambda"
         Height          =   375
         Left            =   3360
         TabIndex        =   17
         Top             =   840
         Width           =   615
      End
      Begin VB.Label LambdaValue 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.TextBox Status 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   6480
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RECORD DATA"
      Height          =   495
      Left            =   10680
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command0 
      Caption         =   "SENSORS OFF"
      Height          =   495
      Left            =   8640
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "MODE1"
         Height          =   495
         Index           =   0
         Left            =   7200
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MODE0"
         Height          =   495
         Index           =   0
         Left            =   5520
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Dyno RPM"
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label DynoRev 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "MODE0"
         Height          =   495
         Index           =   1
         Left            =   5520
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "MODE1"
         Height          =   495
         Index           =   1
         Left            =   7200
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Engine RPM"
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   840
         Width           =   735
      End
      Begin VB.Label EngineRev 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label RecordNo 
      Caption         =   "Record No. : "
      Height          =   255
      Left            =   8640
      TabIndex        =   18
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "DynoTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents comm As MSComm
Attribute comm.VB_VarHelpID = -1
Dim rcvbuffer As String
Dim DrumPeriod(6000) As Double
Dim DrumRPM(6000) As Double
Dim EngineRPM(6000) As Double
Dim Lambda(6000) As Double
Dim RawDyno
Dim RawEngine1
Dim RawEngine2
Dim RawLambda

Dim EngineSum As Double
Dim DrumSum As Double
Dim DataNo As Long
Dim Stroke As Integer
Dim Pulse As Integer
Dim Mode As String


Private Const DatabaseExt = "\OpenDyno.ini"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(Section As String, KeyName As String, Filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), Filename))
End Function

Public Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function






'---OnComm is an event-------
Private Sub comm_OnComm()
    temp = comm.Input
    rcvbuffer = rcvbuffer & temp
    marker = InStr(rcvbuffer, Chr(10))
    rcvtxt = Left$(rcvbuffer, marker + 1)
    rcvbuffer = Mid$(rcvbuffer, marker + 1)
    If InStr(rcvtxt, "DATA") <> 0 Then Status.Text = "DATA"
    If InStr(rcvtxt, "STOP") <> 0 Then Status.Text = "STOP"
    If InStr(rcvtxt, "BT1") <> 0 Then rcvtxt = "b1"
    If InStr(rcvtxt, "BT2") <> 0 Then rcvtxt = "b2"
    If InStr(rcvtxt, "BT3") <> 0 Then rcvtxt = "b3"
    If InStr(rcvtxt, "READY") <> 0 Then Status.Text = "READY"
    If InStr(rcvtxt, ";") And Len(rcvtxt) > 5 Then
        Value = Split(rcvtxt, ";", 5)
        RawDyno = Value(0)
        DynoCalc
        RawEngine1 = Value(1)
        RawEngine2 = Value(2)
        EngineCalc
        RawLambda = Value(3)
        LambdaCalc
        DispData
        If Mode = "record" Then DataNo = DataNo + 1
        If DataNo = 5999 Then SwitchMode "save"
    End If
    
        
End Sub

Public Function EngineCalc()
    If RawEngine2 <> 0 Then
        EngineRPM(DataNo) = Round(60 / ((RawEngine1 / (RawEngine2)) * 0.00005))
        EngineSum = EngineRPM(DataNo)
    Else
        EngineRPM(DataNo) = 0
        EngineSum = 0
    End If
End Function

Public Function DynoCalc()
    DrumSum = DrumSum + ((RawDyno * 0.00005) * Pulse)
    DrumPeriod(DataNo) = DrumSum / 2
    DrumSum = DrumSum - DrumPeriod(DataNo)
    DrumRPM(DataNo) = Round(60 / DrumPeriod(DataNo))
End Function

Public Function LambdaCalc()
    Lambda(DataNo) = RawLambda
End Function




Public Function START() As String
Set comm = New MSComm
            
    

    
            '-------READING COMM PARAMETERS
            comm.Settings = "38400,n,8,2"
            comm.Handshaking = 1
            comm.CommPort = CLng(ReadINI("ComSettings", "ComPort", App.Path & DatabaseExt))
            comm.RThreshold = 1
            comm.InBufferSize = 2048
            comm.OutBufferSize = 2048
            On Error Resume Next
            comm.PortOpen = True
            If comm.PortOpen = False Then START = "NoComm"
            '-----------------------------------
            
End Function

Private Sub DispData()

DynoRev.Caption = DrumRPM(DataNo)
EngineRev.Caption = EngineRPM(DataNo)
LambdaValue.Caption = Lambda(DataNo)
RecordNo.Caption = "Record No. : " & DataNo
End Sub

Private Sub SwitchMode(temp As String)
Mode = temp
Select Case Mode
    Case "standby"
        comm.Output = "MODE0"
        Command0.Caption = "SENSORS ON"
        Command1.Caption = "RECORD DATA"
    Case "sensors"
        comm.Output = "MODE1"
        Command0.Caption = "SENSORS OFF"
        Command1.Caption = "RECORD DATA"
    Case "record"
        DrumSum = 0
        DataNo = 0
        Command0.Caption = "SENSORS OFF"
        Command1.Caption = "RECORDING ..."
        comm.Output = "MODE1"
    Case "save"
        comm.Output = "MODE0"
        Command0.Caption = "SENSORS ON"
        Command1.Caption = "RECORD DATA"
        CommonDialog1.ShowSave
        Open CommonDialog1.Filename For Output As #1
            Print #1, DataNo
            For I = 0 To DataNo
                Print #1, DrumPeriod(I) & ";" & DrumRPM(I) & ";" & EngineRPM(I)
            Next I
        Close #1
End Select
    
End Sub



Private Sub Command0_Click()
Select Case Mode
    Case "sensors"
        SwitchMode ("standby")
    Case "record"
        SwitchMode ("save")
    Case "standby"
        SwitchMode ("sensors")
    Case "save"
        SwitchMode ("sensors")
End Select
End Sub

Private Sub Command1_Click()
If Mode = "record" Then SwitchMode ("save") Else SwitchMode ("record")
End Sub




Private Sub Form_Load()

Pulse = CLng(ReadINI("Dyno", "Pulse", App.Path & DatabaseExt))
Stroke = CLng(ReadINI("Dyno", "Stroke", App.Path & DatabaseExt))

Summe = 0

If START = "NoComm" Then
MsgBox "Com Port Error", 16, "Error"
Unload Me
End If

SwitchMode ("standby")

End Sub


