VERSION 5.00
Begin VB.Form ClimateCorrection 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Climatic Correction"
   ClientHeight    =   3570
   ClientLeft      =   2040
   ClientTop       =   1380
   ClientWidth     =   4575
   Icon            =   "ClimateCorrection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4575
   Begin VB.Frame Frame2 
      Caption         =   "Manual Correction Factor"
      Height          =   735
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   4095
      Begin VB.TextBox CorrectionValue 
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
      Caption         =   "EWG 80/1269"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.TextBox Humidity 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Text            =   "50"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Temperature 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "20"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Pressure 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "990"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "%"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Humidity"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "°C"
         Height          =   255
         Left            =   1560
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Temperature"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "mbar"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Atmospheric Pressure"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "ClimateCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
temp = CInt(Pressure.Text) - (CDbl(Humidity.Text) / 100) * 2.408 * 10 ^ 11 * ((300 / (CDbl(Temperature.Text) + 273.15)) ^ 5) * Exp(-22.644 * (300 / (CDbl(Temperature.Text) + 273.15)))
Main.Correction = ((990 / temp) ^ 1.2) * (((CDbl(Temperature.Text) + 273.15) / 298) ^ 0.6)
Main.CorrectionFactor.Caption = Format(Main.Correction, "0.0000")
Main.CorrectionType = "EWG 80/1269"
Main.Temperature = CDbl(Temperature.Text)
Main.Pressure = CDbl(Pressure.Text)
Main.Humidity = CDbl(Humidity.Text)
Unload ClimateCorrection
End Sub

Private Sub Command2_Click()
Main.Correction = CDbl(CorrectionValue.Text)
Main.CorrectionFactor.Caption = Format(Main.Correction, "0.0000")
If Main.Correction <> 1 Then Main.CorrectionType = "Other" Else Main.CorrectionType = "None"
Unload ClimateCorrection
End Sub
