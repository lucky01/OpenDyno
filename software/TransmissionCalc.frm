VERSION 5.00
Begin VB.Form TransmissionCalc 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Transmission"
   ClientHeight    =   4095
   ClientLeft      =   2040
   ClientTop       =   1380
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4575
   Begin VB.Frame Frame3 
      Caption         =   "Max Engine RPM"
      Height          =   735
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   4095
      Begin VB.TextBox MaxEngineRPM 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Manual Factor"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   4095
      Begin VB.TextBox TransmissionValue 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vehicle Data"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Gear 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "0"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Wheel 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   ": 1"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Gear Transmission Ratio"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "mm"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Wheel Circumference"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "TransmissionCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
TransmissionValue.Text = Round((Main.Diameter * 1000) / (CInt(Wheel) / 3.14159265358979) * CDbl(Gear), 2)
Main.TransmissionValue.Caption = TransmissionValue.Text
Main.Transmission = CDbl(TransmissionValue.Text)
Unload TransmissionCalc
End Sub

Private Sub Command2_Click()
On Error Resume Next
Main.Transmission = CDbl(TransmissionValue.Text)
Main.TransmissionValue.Caption = Main.Transmission
Unload TransmissionCalc
End Sub

Private Sub Command3_Click()
If Main.MaxRPM <> 0 Then
On Error Resume Next
TransmissionValue.Text = Round(CInt(MaxEngineRPM) / Main.MaxRPM, 2)
Main.TransmissionValue.Caption = TransmissionValue.Text
Main.Transmission = CDbl(TransmissionValue.Text)
Unload TransmissionCalc
End If

End Sub

