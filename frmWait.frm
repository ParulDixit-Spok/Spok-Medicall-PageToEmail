VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Wait"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbWait 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please wait for application closing!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   4275
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private d As Single

Private Sub Form_Load()
d = 0
End Sub

Private Sub Timer1_Timer()
d = d + 10
If d > pbWait.Max Then
  pbWait.Value = pbWait.Min
Else
  pbWait.Value = d
End If
End Sub
