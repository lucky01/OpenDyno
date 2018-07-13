VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form DisplayResults 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Dyno Results"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12015
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5895
      Left            =   4440
      OleObjectBlob   =   "DisplayResults.frx":0000
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   600
      Width           =   7335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Preview"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   4080
      TabIndex        =   29
      Top             =   240
      Width           =   7815
      Begin MSComctlLib.Slider ZoomYMin 
         Height          =   2895
         Left            =   120
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   5106
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Max             =   5
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider ZoomYMax 
         Height          =   5895
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   10398
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   -5
         Max             =   0
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider ZoomXMax 
         Height          =   255
         Left            =   3960
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   6360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   10
         Max             =   20
         SelStart        =   10
         TickStyle       =   3
         Value           =   10
      End
      Begin MSComctlLib.Slider ZoomXMin 
         Height          =   255
         Left            =   240
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   6360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         TickStyle       =   3
      End
      Begin VB.CheckBox Check_Power 
         Caption         =   "Power"
         Height          =   375
         Left            =   6840
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   6800
         Width           =   800
      End
      Begin VB.CheckBox Check_RPM 
         Caption         =   "RPM"
         Height          =   375
         Left            =   5280
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   6800
         Width           =   735
      End
      Begin VB.CheckBox Check_Torque 
         Caption         =   "Torque"
         Height          =   375
         Left            =   6000
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   6800
         Width           =   855
      End
      Begin VB.Frame Hide 
         Caption         =   "Hide Graph"
         Height          =   615
         Left            =   5160
         TabIndex        =   36
         Top             =   6600
         Width           =   2535
      End
      Begin VB.ComboBox CompareSelection 
         Height          =   315
         ItemData        =   "DisplayResults.frx":2385
         Left            =   2880
         List            =   "DisplayResults.frx":238F
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   6720
         Width           =   2175
      End
      Begin VB.ComboBox DataSelection 
         Height          =   315
         ItemData        =   "DisplayResults.frx":23A6
         Left            =   120
         List            =   "DisplayResults.frx":23B3
         Style           =   2  'Dropdown-Liste
         TabIndex        =   4
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Compare with"
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   6770
         Width           =   975
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   135
         Left            =   120
         Top             =   7080
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Engine 
      Caption         =   "Engine Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   3855
      Begin VB.Label Note 
         Caption         =   "Note: Engine RPM calculated from Drum RPM !"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2760
         Width           =   3495
      End
      Begin VB.Label Label10 
         Caption         =   "at RPM"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "at RPM"
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
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Max RPM:"
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
         Left            =   240
         TabIndex        =   22
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Torque (Nm):"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Power (kW):"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label ERPM 
         Caption         =   "10000"
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
         Left            =   1920
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label ETorque 
         Caption         =   "00,00"
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
         Left            =   1920
         TabIndex        =   18
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label ETorqueRPM 
         Caption         =   "10000"
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
         Left            =   1920
         TabIndex        =   17
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label EPowerRPM 
         Caption         =   "10000"
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
         Left            =   1920
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label EPower 
         Caption         =   "00,00"
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
         Left            =   1920
         TabIndex        =   15
         Top             =   1920
         Width           =   1815
      End
   End
   Begin VB.Frame Wheel 
      Caption         =   "Wheel Data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.Label AccelTime 
         Caption         =   "00.00"
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
         Left            =   1920
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "in Seconds:"
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
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "at km/h"
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
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "at km/h"
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
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label WPower 
         Caption         =   "00,00"
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
         Left            =   1920
         TabIndex        =   13
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label WPowerSpeed 
         Caption         =   "100.0"
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
         Left            =   1920
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label WTorqueSpeed 
         Caption         =   "100.0"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label WTorque 
         Caption         =   "00,00"
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
         Left            =   1920
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label WSpeed 
         Caption         =   "100.0"
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
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Power (kW):"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Torque Nm:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Speed (km/h):"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   135
      Left            =   1440
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Ausgefüllt
      Height          =   135
      Left            =   120
      Top             =   7440
      Width           =   1215
   End
End
Attribute VB_Name = "DisplayResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check_Power_Click()
    Call Main.DrawChart
End Sub

Private Sub Check_RPM_Click()
    Call Main.DrawChart
End Sub

Private Sub Check_Torque_Click()
    Call Main.DrawChart
End Sub

Private Sub Command1_Click()
    Main.Remote = False
    Main.CommonDialog1.Filename = "dyno" & CStr(Format(Main.Run, "000"))
    Call Main.SaveData
End Sub

Private Sub Command2_Click()
    Main.Remote = False
    Call Main.PrintChart
End Sub

Private Sub Command4_Click()
    Call Main.LoadData_Click
End Sub



Private Sub CompareSelection_Click()

If CompareSelection.ListIndex = 0 Then
    Call Main.ResetChart
Else
    Call Main.CompareChart
End If

End Sub

Private Sub DataSelection_Click()
Select Case DataSelection.ListIndex
    Case 0
          ZoomYMax.Height = 2895
          ZoomYMax.Min = -10
          ZoomYMin.Visible = True
          ZoomXMin.Value = Main.SpeedXMin / 5
          ZoomXMax.Value = Main.SpeedXMax / 5
          ZoomYMin.Value = Main.GraphYMin / -10
          ZoomYMax.Value = Main.GraphYMax / -10
    Case 1
          ZoomYMax.Height = 5895
          ZoomYMax.Min = -25
          ZoomYMin.Visible = False
          ZoomXMin.Value = Main.SpeedXMin / 5
          ZoomXMax.Value = Main.SpeedXMax / 5
          ZoomYMax.Value = Main.EngGraphYMax * -1
    Case 2
          ZoomYMax.Height = 5895
          ZoomYMax.Min = -25
          ZoomYMin.Visible = False
          ZoomYMax.Value = Main.EngGraphYMax * -1
          ZoomXMin.Value = Main.RPMXMin
          ZoomXMax.Value = Main.RPMXMax
End Select
Call Main.DrawChart
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Main.SwitchMode("sensors")
End Sub




Private Sub MSChart1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Part As Integer, Series As Integer, DataPoint As Integer, index3 As Integer, index4 As Integer
Dim valueX As Double, valueY As Double, nullflag As Integer, AxisText As String
    
    With MSChart1
        'Obtain the part of the chart the mouse is on
        .TwipsToChartPart x, y, Part, Series, DataPoint, index3, index4
        If Part = VtChPartTypePoint Then    'Is it a dataPoint?
            .DataGrid.GetData DataPoint, Series + 1, valueY, nullflag  'Get Y value
            .DataGrid.GetData DataPoint, Series, valueX, nullflag  'Get X value
            If .Plot.Axis(VtChAxisIdX).AxisTitle.Text = "1/min x 1000" Then
                    valueX = valueX * 1000
                    AxisText = "rpm"
            Else
                    AxisText = .Plot.Axis(VtChAxisIdX).AxisTitle.Text
            End If
            
            Select Case Series
                Case 1
                    .ToolTipText = Format(valueY, "0.00") & " Nm @ " & Format(valueX, "0") & " " & AxisText
                Case 3
                    .ToolTipText = Format(valueY, "0.00") & " kW " & " (" & Format(valueY * 1.359622, "0.00") & " PS) @ " & Format(valueX, "0") & " " & AxisText
                Case 5
                    .ToolTipText = Format(valueY * 1000, "0") & " rpm @ " & Format(valueX, "0") & " " & AxisText
                Case 7
                    .ToolTipText = Format(valueY, "0.00") & " Nm @ " & Format(valueX, "0") & " " & AxisText
                Case 9
                    .ToolTipText = Format(valueY, "0.00") & " kW " & " (" & Format(valueY * 1.359622, "0.00") & " PS) @ " & Format(valueX, "0") & " " & AxisText
                Case 11
                    .ToolTipText = Round(valueY * 1000, 0) & " rpm @ " & Format(valueX, "0") & " " & AxisText
            End Select
        Else
            .ToolTipText = ""
        End If
    End With
End Sub

Private Sub MSChart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
MSChart1.AllowSelections = True

Dim U As Integer
U = DataPoint
x = Series
Value = MSChart1.ChartData(DataPoint, Series + 1)

    MSChart1.ToolTipText = Round(Value, 2)
    
End Sub


Private Sub ZoomXMin_scroll()
Select Case DataSelection.ListIndex
    Case 0
        Main.SpeedXMin = ZoomXMin.Value * 5
        Main.SpeedXDiv = (Main.SpeedXMax - Main.SpeedXMin) / 5
        Call Main.DrawChart
        ZoomXMin.Text = "min " & Main.SpeedXMin & " km/h"
    Case 1
        Main.SpeedXMin = ZoomXMin.Value * 5
        Main.SpeedXDiv = (Main.SpeedXMax - Main.SpeedXMin) / 5
        Call Main.DrawChart
        ZoomXMin.Text = "min " & Main.SpeedXMin & " km/h"
    Case 2
        Main.RPMXMin = ZoomXMin.Value
        Main.RPMXDiv = Main.RPMXMax - Main.RPMXMin
        Call Main.DrawChart
        ZoomXMin.Text = "min " & Main.RPMXMin * 1000 & " rpm"
End Select
End Sub

Private Sub ZoomXMax_scroll()
Select Case DataSelection.ListIndex
    Case 0
        Main.SpeedXMax = ZoomXMax.Value * 5
        Main.SpeedXDiv = (Main.SpeedXMax - Main.SpeedXMin) / 5
        Call Main.DrawChart
        ZoomXMax.Text = "max " & Main.SpeedXMax & " km/h"
    Case 1
        Main.SpeedXMax = ZoomXMax.Value * 5
        Main.SpeedXDiv = (Main.SpeedXMax - Main.SpeedXMin) / 5
        Call Main.DrawChart
        ZoomXMax.Text = "max " & Main.SpeedXMax & " km/h"
    Case 2
        Main.RPMXMax = ZoomXMax.Value
        Main.RPMXDiv = Main.RPMXMax - Main.RPMXMin
        Call Main.DrawChart
        ZoomXMax.Text = "max " & Main.RPMXMax * 1000 & " rpm"
End Select
End Sub

Private Sub ZoomYMin_scroll()
Select Case DataSelection.ListIndex
    Case 0
        Main.GraphYMin = ZoomYMin.Value * -10
        Main.GraphYDiv = (Main.GraphYMax - Main.GraphYMin) / 5
        Call Main.DrawChart
        ZoomYMin.Text = "min " & Main.GraphYMin & " Nm/kW"
End Select
End Sub

Private Sub ZoomYMax_scroll()
Select Case DataSelection.ListIndex
    Case 0
        Main.GraphYMax = ZoomYMax.Value * -10
        Main.GraphYDiv = (Main.GraphYMax - Main.GraphYMin) / 5
        Call Main.DrawChart
        ZoomYMax.Text = "max " & Main.GraphYMax & " Nm/kW"
    Case 1
        Main.EngGraphYMax = ZoomYMax.Value * -1
        Main.EngGraphYDiv = Main.EngGraphYMax
        Call Main.DrawChart
        ZoomYMax.Text = "max " & Main.EngGraphYMax & " Nm/kW"
    Case 2
        Main.EngGraphYMax = ZoomYMax.Value * -1
        Main.EngGraphYDiv = Main.EngGraphYMax
        Call Main.DrawChart
        ZoomYMax.Text = "max " & Main.EngGraphYMax & " Nm/kW"
End Select
End Sub
