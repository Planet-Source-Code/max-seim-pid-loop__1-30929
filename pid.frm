VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Simple PID Loop Simulator - Industrial Level Control ... mlseim@mmm.com"
   ClientHeight    =   7020
   ClientLeft      =   3090
   ClientTop       =   2715
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   9420
   Begin VB.CommandButton Command3 
      Caption         =   "Create unstable water supply"
      Height          =   435
      Left            =   90
      TabIndex        =   46
      Top             =   5790
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   1125
      Left            =   5190
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   45
      Top             =   5850
      Width           =   4125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1305
      Left            =   5160
      ScaleHeight     =   83
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   44
      Top             =   4080
      Width           =   4125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto Control"
      Height          =   315
      Left            =   5100
      TabIndex        =   27
      Top             =   540
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Manual Control"
      Height          =   315
      Left            =   5100
      TabIndex        =   26
      Top             =   180
      Width           =   1245
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      TabIndex        =   20
      Text            =   "10"
      Top             =   1800
      Width           =   645
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      TabIndex        =   19
      Text            =   "3"
      Top             =   1410
      Width           =   645
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      TabIndex        =   18
      Text            =   "30"
      Top             =   990
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   90
      Top             =   3810
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   660
      TabIndex        =   11
      Top             =   2670
      Width           =   675
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3825
      Left            =   4020
      Max             =   0
      Min             =   3100
      TabIndex        =   7
      Top             =   690
      Width           =   255
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   810
      Max             =   100
      TabIndex        =   4
      Top             =   4650
      Width           =   1365
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   90
      Max             =   100
      TabIndex        =   1
      Top             =   930
      Width           =   1425
   End
   Begin VB.Label Label39 
      Caption         =   $"pid.frx":0000
      ForeColor       =   &H00808000&
      Height          =   675
      Left            =   120
      TabIndex        =   54
      Top             =   6360
      Width           =   4545
   End
   Begin VB.Label Label15 
      Caption         =   $"pid.frx":009F
      ForeColor       =   &H00808000&
      Height          =   795
      Left            =   6840
      TabIndex        =   53
      Top             =   150
      Width           =   2535
   End
   Begin VB.Line Line39 
      BorderWidth     =   2
      X1              =   30
      X2              =   9330
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Process Variable - PV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6210
      TabIndex        =   52
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label42 
      Caption         =   "Output Valve Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6240
      TabIndex        =   51
      Top             =   5610
      Width           =   1905
   End
   Begin VB.Label Label41 
      Caption         =   "100%"
      Height          =   285
      Left            =   4740
      TabIndex        =   50
      Top             =   5850
      Width           =   435
   End
   Begin VB.Label Label40 
      Caption         =   "0%"
      Height          =   315
      Left            =   4890
      TabIndex        =   49
      Top             =   6750
      Width           =   285
   End
   Begin VB.Label Label38 
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   60
      TabIndex        =   48
      Top             =   1440
      Width           =   435
   End
   Begin VB.Label Label37 
      Caption         =   "% Valve Position"
      Height          =   255
      Left            =   2790
      TabIndex        =   47
      Top             =   5310
      Width           =   1245
   End
   Begin VB.Shape Shape19 
      FillColor       =   &H0000FF00&
      Height          =   285
      Left            =   1380
      Shape           =   3  'Circle
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label36 
      Caption         =   "Rate"
      Height          =   165
      Left            =   6750
      TabIndex        =   43
      Top             =   3270
      Width           =   375
   End
   Begin VB.Label Label35 
      Caption         =   "d"
      Height          =   195
      Left            =   6540
      TabIndex        =   42
      Top             =   3150
      Width           =   135
   End
   Begin VB.Line Line38 
      X1              =   6570
      X2              =   6570
      Y1              =   3120
      Y2              =   2880
   End
   Begin VB.Line Line37 
      X1              =   6960
      X2              =   7050
      Y1              =   3120
      Y2              =   3180
   End
   Begin VB.Line Line36 
      X1              =   6960
      X2              =   6870
      Y1              =   3120
      Y2              =   3180
   End
   Begin VB.Line Line35 
      X1              =   6960
      X2              =   6960
      Y1              =   3240
      Y2              =   3120
   End
   Begin VB.Line Line34 
      X1              =   6720
      X2              =   6960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape Shape18 
      Height          =   255
      Left            =   6510
      Top             =   3120
      Width           =   225
   End
   Begin VB.Label Label34 
      Caption         =   "SP - PV"
      Height          =   225
      Left            =   6660
      TabIndex        =   41
      Top             =   2250
      Width           =   585
   End
   Begin VB.Label Label33 
      Caption         =   "0-120 Seconds"
      Height          =   225
      Left            =   6660
      TabIndex        =   40
      Top             =   1860
      Width           =   1245
   End
   Begin VB.Label Label32 
      Caption         =   "Reset"
      Height          =   225
      Left            =   7170
      TabIndex        =   39
      Top             =   2610
      Width           =   465
   End
   Begin VB.Line Line33 
      X1              =   6960
      X2              =   6870
      Y1              =   2880
      Y2              =   2970
   End
   Begin VB.Line Line32 
      X1              =   6870
      X2              =   6960
      Y1              =   2790
      Y2              =   2880
   End
   Begin VB.Line Line31 
      X1              =   6870
      X2              =   7020
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line30 
      X1              =   7020
      X2              =   6870
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label Label31 
      Caption         =   "OUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7620
      TabIndex        =   38
      Top             =   3120
      Width           =   405
   End
   Begin VB.Line Line29 
      X1              =   6360
      X2              =   6450
      Y1              =   3540
      Y2              =   3600
   End
   Begin VB.Line Line28 
      X1              =   6360
      X2              =   6450
      Y1              =   3540
      Y2              =   3480
   End
   Begin VB.Line Line27 
      X1              =   6360
      X2              =   7500
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Line Line26 
      X1              =   7500
      X2              =   7500
      Y1              =   2880
      Y2              =   3540
   End
   Begin VB.Line Line25 
      X1              =   7140
      X2              =   7500
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label30 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6030
      TabIndex        =   37
      Top             =   2220
      Width           =   585
   End
   Begin VB.Label Label29 
      Caption         =   "e = error (difference)"
      Height          =   225
      Left            =   4560
      TabIndex        =   36
      Top             =   2220
      Width           =   1425
   End
   Begin VB.Label Label28 
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   5400
      TabIndex        =   35
      Top             =   2670
      Width           =   195
   End
   Begin VB.Line Line22 
      X1              =   5670
      X2              =   5610
      Y1              =   2880
      Y2              =   2970
   End
   Begin VB.Line Line21 
      X1              =   5670
      X2              =   5610
      Y1              =   2880
      Y2              =   2820
   End
   Begin VB.Line Line24 
      X1              =   6690
      X2              =   6630
      Y1              =   2880
      Y2              =   2970
   End
   Begin VB.Line Line23 
      X1              =   6690
      X2              =   6630
      Y1              =   2880
      Y2              =   2790
   End
   Begin VB.Line Line20 
      X1              =   5070
      X2              =   5130
      Y1              =   3120
      Y2              =   3210
   End
   Begin VB.Line Line19 
      X1              =   5070
      X2              =   5010
      Y1              =   3120
      Y2              =   3210
   End
   Begin VB.Line Line18 
      X1              =   4830
      X2              =   4740
      Y1              =   2880
      Y2              =   2940
   End
   Begin VB.Line Line17 
      X1              =   4830
      X2              =   4740
      Y1              =   2880
      Y2              =   2820
   End
   Begin VB.Label Label27 
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4440
      TabIndex        =   34
      Top             =   2640
      Width           =   315
   End
   Begin VB.Label Label26 
      Caption         =   "PV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5250
      TabIndex        =   33
      Top             =   3300
      Width           =   315
   End
   Begin VB.Line Line16 
      X1              =   5070
      X2              =   5580
      Y1              =   3540
      Y2              =   3540
   End
   Begin VB.Line Line15 
      X1              =   5070
      X2              =   5070
      Y1              =   3120
      Y2              =   3540
   End
   Begin VB.Label Label25 
      Caption         =   "Process"
      Height          =   225
      Left            =   5670
      TabIndex        =   32
      Top             =   3420
      Width           =   645
   End
   Begin VB.Label Label24 
      Caption         =   "X Gain"
      Height          =   225
      Left            =   5730
      TabIndex        =   31
      Top             =   2760
      Width           =   525
   End
   Begin VB.Line Line14 
      X1              =   6300
      X2              =   6690
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line13 
      X1              =   5310
      X2              =   5670
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line12 
      X1              =   4830
      X2              =   4500
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line11 
      X1              =   4950
      X2              =   5160
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line10 
      X1              =   5040
      X2              =   5160
      Y1              =   2760
      Y2              =   2970
   End
   Begin VB.Line Line9 
      X1              =   5040
      X2              =   4950
      Y1              =   2790
      Y2              =   2970
   End
   Begin VB.Shape Shape17 
      Height          =   405
      Left            =   5580
      Top             =   3330
      Width           =   795
   End
   Begin VB.Shape Shape16 
      Height          =   495
      Left            =   6690
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   465
   End
   Begin VB.Shape Shape15 
      Height          =   435
      Left            =   5670
      Top             =   2670
      Width           =   645
   End
   Begin VB.Shape Shape14 
      Height          =   495
      Left            =   4830
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   465
   End
   Begin VB.Label Label23 
      Caption         =   "OUT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2790
      TabIndex        =   30
      Top             =   5010
      Width           =   645
   End
   Begin VB.Label Label22 
      Caption         =   "SP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   4320
      TabIndex        =   29
      Top             =   4170
      Width           =   435
   End
   Begin VB.Label Label21 
      Caption         =   "PV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   210
      TabIndex        =   28
      Top             =   2640
      Width           =   435
   End
   Begin VB.Shape Shape13 
      FillColor       =   &H0000FF00&
      Height          =   285
      Left            =   6390
      Shape           =   3  'Circle
      Top             =   570
      Width           =   255
   End
   Begin VB.Shape Shape12 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   6390
      Shape           =   3  'Circle
      Top             =   210
      Width           =   255
   End
   Begin VB.Label Label20 
      Caption         =   "0-120 Seconds"
      Height          =   225
      Left            =   6660
      TabIndex        =   25
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label19 
      Caption         =   "0-100%"
      Height          =   225
      Left            =   6660
      TabIndex        =   24
      Top             =   1080
      Width           =   585
   End
   Begin VB.Label Label18 
      Caption         =   "Derivitave (RATE)"
      Height          =   255
      Left            =   4680
      TabIndex        =   23
      Top             =   1830
      Width           =   1305
   End
   Begin VB.Label Label17 
      Caption         =   "Integral (RESET)"
      Height          =   285
      Left            =   4740
      TabIndex        =   22
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Label Label16 
      Caption         =   "Proportional (GAIN)"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   1050
      Width           =   1395
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H00FFFF00&
      Height          =   405
      Left            =   2460
      Top             =   5010
      Width           =   285
   End
   Begin VB.Shape Shape10 
      FillColor       =   &H00FFFF00&
      Height          =   225
      Left            =   2460
      Top             =   4170
      Width           =   285
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      Height          =   3075
      Left            =   2190
      Top             =   1050
      Width           =   225
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00FFFF00&
      Height          =   435
      Left            =   2190
      Top             =   600
      Width           =   195
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1350
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label14 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2370
      TabIndex        =   17
      Top             =   4560
      Width           =   465
   End
   Begin VB.Label Label13 
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
      Left            =   750
      TabIndex        =   16
      Top             =   420
      Width           =   465
   End
   Begin VB.Label Label12 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3810
      TabIndex        =   15
      Top             =   4530
      Width           =   555
   End
   Begin VB.Label Label11 
      Caption         =   "3000 Gal/Min Max."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   780
      TabIndex        =   14
      Top             =   5130
      Width           =   1425
   End
   Begin VB.Label Label10 
      Caption         =   " Gal/Min Max."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label9 
      Caption         =   "Actual Level"
      Height          =   405
      Left            =   750
      TabIndex        =   12
      Top             =   2970
      Width           =   525
   End
   Begin VB.Label Label8 
      Caption         =   "Tank"
      Height          =   225
      Left            =   780
      TabIndex        =   10
      Top             =   2400
      Width           =   435
   End
   Begin VB.Label Label7 
      Caption         =   "3100 Gallon"
      Height          =   255
      Left            =   570
      TabIndex        =   9
      Top             =   2190
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   " Gallons"
      Height          =   225
      Left            =   4320
      TabIndex        =   8
      Top             =   4560
      Width           =   645
   End
   Begin VB.Label Label5 
      Caption         =   "Level Setpoint"
      Height          =   225
      Left            =   3630
      TabIndex        =   6
      Top             =   390
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   3135
      Left            =   3690
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label Label4 
      Caption         =   "Position 0-100%"
      Height          =   225
      Left            =   900
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Valve"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   4320
      Width           =   465
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   2
      Height          =   435
      Left            =   2460
      Top             =   5010
      Width           =   315
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   2
      Height          =   255
      Left            =   2460
      Top             =   4140
      Width           =   315
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   990
      Y2              =   420
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2190
      X2              =   2190
      Y1              =   630
      Y2              =   1020
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1350
      X2              =   2190
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1350
      X2              =   2400
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   615
      Index           =   1
      Left            =   2220
      Top             =   4380
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Position 0-100%"
      Height          =   255
      Left            =   210
      TabIndex        =   2
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Manual Valve"
      Height          =   225
      Left            =   1380
      TabIndex        =   0
      Top             =   150
      Width           =   1065
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   615
      Index           =   0
      Left            =   600
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   420
      Width           =   585
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   3100
      Index           =   1
      Left            =   3690
      Top             =   1020
      Width           =   165
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C000&
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   3015
      Index           =   0
      Left            =   1620
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Line Line4 
      BorderWidth     =   4
      X1              =   3660
      X2              =   1590
      Y1              =   4140
      Y2              =   4140
   End
   Begin VB.Line Line3 
      BorderWidth     =   4
      X1              =   3660
      X2              =   3660
      Y1              =   1050
      Y2              =   4140
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   1590
      X2              =   3660
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   1590
      X2              =   1590
      Y1              =   1020
      Y2              =   4140
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnInstructions 
      Caption         =   "Instructions"
   End
   Begin VB.Menu mnAbout 
      Caption         =   "About"
      Begin VB.Menu mnLine1 
         Caption         =   "By Max Seim - mlseim@mmm.com"
      End
      Begin VB.Menu mnLine2 
         Caption         =   "Simple PID loop simulator - Version 01.18.02"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'  Simple PID Loop Simulator ... for educational use.
'
'  By Max Seim,  mlseim@mmm.com
'     Systems Control Technician, 3M Company, Cottage Grove, Minnesota
'
Dim invalve As Integer
Dim outvalve As Integer
Dim mode As Integer '0=manual, 1=auto
Dim error As Integer
Dim stability As Integer '0=stable, 1=unstable
Dim supply As Integer
Dim x As Integer
Dim y As Integer
Dim gain As Long
Dim reset As Long
Dim rate As Long
Dim output As Integer
Dim pv As Long
Dim s1 As Integer
Dim s2 As Integer
Dim outgraph(100) As Integer
Dim pvgraph(100) As Integer
Dim inputd As Long
Dim inputdf As Long
Dim inputlast As Long
Dim feedback As Long
Dim dfilter As Long
Dim sp As Long

Private Sub Command1_Click() ' Go into MANUAL control
Shape12.FillStyle = 0
Shape13.FillStyle = 1
mode = 0
HScroll2.Enabled = True
End Sub

Private Sub Command2_Click() ' Go into AUTO control
Shape12.FillStyle = 1
Shape13.FillStyle = 0
mode = 1
HScroll2.Enabled = False
End Sub

Private Sub Command3_Click() ' Toggle the Unstable Water Supply
'
If stability = 0 Then
   Shape19.FillStyle = 0
   stability = 1
   Exit Sub
End If
If stability = 1 Then
   Shape19.FillStyle = 1
   stability = 0
   supply = 2000
End If
Label38.Caption = supply
End Sub

Private Sub Form_load()
' center the form
Form1.Left = (Screen.Width / 2) - (Form1.Width / 2)
Form1.Top = (Screen.Height / 2) - (Form1.Height / 2)

' Initialize the Sliders and other variables
VScroll1.Value = 0
Shape1(1).Top = (3100 - VScroll1.Value) + 1040
Shape1(1).Height = VScroll1.Value
Shape1(0).Top = (3100 - pv) + 1030
Shape1(0).Height = pv
Label12.Caption = VScroll1.Value

supply = 2000
Label38.Caption = supply
HScroll1.Value = 100
Label13.Caption = HScroll1.Value
invalve = (HScroll1.Value * (supply / 100)) / 60

outvalve = 0
Label14.Caption = 0
Text1.Text = pv

   Shape12.FillStyle = 1
   Shape13.FillStyle = 0
   mode = 1
   HScroll2.Enabled = False

gain = Text2.Text
reset = Text3.Text
rate = Text4.Text
VScroll1.Value = 1500
Shape1(1).Top = (3100 - VScroll1.Value) + 1040
Shape1(1).Height = VScroll1.Value
Label12.Caption = VScroll1.Value
sp = VScroll1.Value

Picture1.Cls
Picture1.ScaleMode = 3
Picture1.ScaleHeight = 3105
Picture1.ScaleWidth = 100
Picture1.AutoRedraw = True
Picture1.ForeColor = vbCyan
Picture1.DrawStyle = 0
Picture1.DrawWidth = 2

Picture2.Cls
Picture2.ScaleMode = 3
Picture2.ScaleHeight = 105
Picture2.ScaleWidth = 100
Picture2.AutoRedraw = True
Picture2.ForeColor = vbRed
Picture2.DrawStyle = 0
Picture2.DrawWidth = 2
End Sub

Private Sub HScroll1_Change() ' Slider control for INLET VALVE POSITION
Label13.Caption = HScroll1.Value
invalve = (HScroll1.Value * (supply / 100)) / 60
End Sub

Private Sub HScroll2_Change() ' Slider control for OUTLET VALVE POSITION
Label14.Caption = HScroll2.Value
outvalve = (HScroll2.Value * 30) / 60
End Sub

Private Sub mnExit_Click() ' Exit program
Erase outgraph
Erase pvgraph
Unload Me
End Sub

Private Sub mnInstructions_Click() ' Show instruction form
instructions.Show
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
   If Text2 < 0 Then
   Text2 = 0
   End If
      If Text2 > 100 Then
      Text2 = 100
      End If
gain = Val(Text2)
   If Text3 < 0 Then
   Text3 = 0
   End If
      If Text3 > 120 Then
      Text3 = 120
      End If
reset = Val(Text3)
   If Text4 < 0 Then
   Text4 = 0
   End If
      If Text4 > 120 Then
      Text4 = 120
      End If
rate = Val(Text4)

  If pv < 3101 Then
     pv = pv + invalve
  End If
     If pv > 0 Then
       pv = pv - outvalve
     End If
Text1.Text = pv
error = sp - pv
Label30.Caption = error
Label38.Caption = supply
tank
If stability = 1 Then
watersupply
   invalve = (HScroll1.Value * (supply / 100)) / 60
End If
  If mode = 1 Then
     pidloop
  End If

' Graph the PV, Process Variable
Picture1.Cls
pvgraph(100) = pv
For x = 0 To 99
pvgraph(x) = pvgraph(x + 1)
  Picture1.PSet (x, 3000 - (pvgraph(x)))
Next x

' Display the SP line (yellow)
Picture1.Line (0, 3000 - sp)-(100, 3000 - sp), vbYellow

' Graph the OUTPUT VALVE position
Picture2.Cls
outgraph(100) = outvalve
For x = 0 To 99
outgraph(x) = outgraph(x + 1)
  Picture2.PSet (x, 100 - (outgraph(x) * 2))
Next x
End Sub

Private Sub VScroll1_Change() ' Slider control for SP (setpoint)
Shape1(1).Top = (3100 - VScroll1.Value) + 1040
Shape1(1).Height = VScroll1.Value
Label12.Caption = VScroll1.Value
sp = VScroll1.Value
End Sub
Private Sub tank() ' Draw the water tank animation
If HScroll1.Value > 0 Then
   Shape7.FillStyle = 0
   Shape8.FillStyle = 0
   Shape9.FillStyle = 0
  Else: Shape7.FillStyle = 1
        Shape8.FillStyle = 1
        Shape9.FillStyle = 1
End If
If pv > -1 Then
   Shape1(0).Top = (3100 - pv) + 1030
   Shape1(0).Height = pv
End If
  If (pv > 0) Or (HScroll1.Value > 0) Then
  Shape10.FillStyle = 0
    Else: Shape10.FillStyle = 1
  End If
     If (HScroll2.Value > 0) And (pv > 0) Then
     Shape11.FillStyle = 0
       Else: Shape11.FillStyle = 1
     End If
End Sub
Private Sub watersupply() ' Create an unstable water supply
  Randomize
  s1 = Int(Rnd(1) * 20 + 1)
     Randomize
     s2 = Int(Rnd(1) * 1000 + 1)
If s2 < 100 Then
  supply = supply + s1
End If
If s2 > 900 Then
  supply = supply - s1
End If
   If supply < 500 Then
   supply = 500
   End If
      If supply > 2500 Then
      supply = 2500
      End If
      
End Sub

Private Sub pidloop()
dfilter = 10 ' Filter value to scale down derivative effect.
inputd = pv + (inputlast - pv) * (rate / 60)
inputlast = pv
inputdf = inputdf + (inputd - inputdf) * dfilter / 60
output = (sp - inputdf) * (gain / 100) + feedback
If output > 100 Then ' clamp output valve between 0 and 100%
  output = 100
End If
If output < 0 Then
  output = 0
End If
HScroll2.Value = 100 - output ' Change slider value (AUTO MODE)
Label14.Caption = HScroll2.Value
feedback = feedback - (feedback - output) * reset / 60
End Sub

