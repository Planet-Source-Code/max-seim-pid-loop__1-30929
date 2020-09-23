VERSION 5.00
Begin VB.Form instructions 
   Caption         =   "Simple PID Loop Explanation and Instructions"
   ClientHeight    =   7260
   ClientLeft      =   4515
   ClientTop       =   2025
   ClientWidth     =   9390
   LinkTopic       =   "Form2"
   ScaleHeight     =   7260
   ScaleWidth      =   9390
   Begin VB.Label Label17 
      Caption         =   $"instructions.frx":0000
      Height          =   435
      Left            =   600
      TabIndex        =   16
      Top             =   6810
      Width           =   7935
   End
   Begin VB.Label Label16 
      Caption         =   $"instructions.frx":00D1
      Height          =   435
      Left            =   780
      TabIndex        =   15
      Top             =   6360
      Width           =   7755
   End
   Begin VB.Label Label15 
      Caption         =   "Proportional (GAIN):  0-100%, how much to amplify the output based on the error (SP-PV)"
      Height          =   255
      Left            =   510
      TabIndex        =   14
      Top             =   6090
      Width           =   7875
   End
   Begin VB.Label Label14 
      Caption         =   $"instructions.frx":0194
      Height          =   465
      Left            =   900
      TabIndex        =   13
      Top             =   5460
      Width           =   7425
   End
   Begin VB.Label Label13 
      Caption         =   "4)  Make changes to the level setpoint (SP) and watch the PID loop control to the new SP."
      Height          =   285
      Left            =   660
      TabIndex        =   12
      Top             =   5250
      Width           =   6795
   End
   Begin VB.Label Label12 
      Caption         =   "See how well you can manually control the level with an unstable water supply."
      Height          =   315
      Left            =   870
      TabIndex        =   11
      Top             =   4920
      Width           =   6795
   End
   Begin VB.Label Label11 
      Caption         =   "Now you have the ability to control both the inlet valve AND the outlet valve."
      Height          =   255
      Left            =   870
      TabIndex        =   10
      Top             =   4710
      Width           =   6135
   End
   Begin VB.Label Label10 
      Caption         =   "3)  Put the system in MANUAL CONTROL."
      Height          =   225
      Left            =   630
      TabIndex        =   9
      Top             =   4500
      Width           =   7335
   End
   Begin VB.Label Label9 
      Caption         =   "Watch the PID loop control make slight adjustments to correct the unstable supply."
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   4140
      Width           =   6525
   End
   Begin VB.Label Label8 
      Caption         =   "This button will cause the main supply water to fluctuate, like a real process normally does."
      Height          =   345
      Left            =   840
      TabIndex        =   7
      Top             =   3930
      Width           =   6915
   End
   Begin VB.Label Label7 
      Caption         =   "2)  With the manual input valve fully opened (100%), click the button for unstable water supply. "
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   3720
      Width           =   7845
   End
   Begin VB.Label Label6 
      Caption         =   "1)  Alter the manual input valve and watch the PID loop in action."
      Height          =   315
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   7605
   End
   Begin VB.Label Label5 
      Caption         =   $"instructions.frx":0244
      Height          =   705
      Left            =   600
      TabIndex        =   4
      Top             =   2580
      Width           =   7245
   End
   Begin VB.Label Label4 
      Caption         =   $"instructions.frx":034A
      Height          =   645
      Left            =   630
      TabIndex        =   3
      Top             =   1740
      Width           =   6825
   End
   Begin VB.Label Label3 
      Caption         =   "Explanation of Proportional, Integral, Derivative (PID) control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   2
      Top             =   30
      Width           =   8085
   End
   Begin VB.Label Label2 
      Caption         =   $"instructions.frx":0456
      Height          =   645
      Left            =   630
      TabIndex        =   1
      Top             =   1080
      Width           =   6075
   End
   Begin VB.Label Label1 
      Caption         =   $"instructions.frx":051D
      Height          =   465
      Left            =   600
      TabIndex        =   0
      Top             =   570
      Width           =   5925
   End
End
Attribute VB_Name = "instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_load()
instructions.Left = (Screen.Width / 2) - (Form1.Width / 2)
instructions.Top = (Screen.Height / 2) - (Form1.Height / 2)
End Sub

