VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form ChartPrinterSheet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Dyno Data"
   ClientHeight    =   9270
   ClientLeft      =   -255
   ClientTop       =   495
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9270
   ScaleMode       =   0  'User
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   9015
      Left            =   0
      OleObjectBlob   =   "PrintChart.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "ChartPrinterSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
