VERSION 5.00
Begin VB.Form frmDisplay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send e-mails from Paging Queue"
   ClientHeight    =   5115
   ClientLeft      =   450
   ClientTop       =   1095
   ClientWidth     =   12435
   Icon            =   "Display.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   12435
   Begin VB.CommandButton comAbout 
      Caption         =   "About"
      Height          =   420
      Left            =   11355
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1150
      Width           =   975
   End
   Begin VB.CommandButton comExit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   11355
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1695
      Width           =   975
   End
   Begin VB.CommandButton comAlarm 
      Caption         =   "Sys Alarms "
      Height          =   420
      Left            =   11355
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   605
      Width           =   975
   End
   Begin VB.Timer GlobalTimer 
      Left            =   225
      Top             =   1800
   End
   Begin VB.Timer TimeWait 
      Enabled         =   0   'False
      Left            =   225
      Top             =   2295
   End
   Begin VB.CommandButton seeErrors 
      Caption         =   "View Log"
      Height          =   420
      Left            =   11355
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   960
   End
   Begin VB.ListBox lstProcess 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   45
      Width           =   11145
   End
   Begin VB.Label lblTimerState 
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   10185
      TabIndex        =   3
      Top             =   4710
      Width           =   1095
   End
   Begin VB.Label lblWhatsUp 
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   4725
      Width           =   9840
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents CallBack As clsCallBack
Attribute CallBack.VB_VarHelpID = -1


Private pSavedActionReminder() As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Public Sub AddToProcessList(ByVal parInfo As String)

With lstProcess
    .AddItem Now & " " & parInfo
    If .ListCount > 100 Then .RemoveItem 0
    .ListIndex = .ListCount - 1
    
End With
End Sub

Private Sub CallBack_Sent(sProfile As String, lErrNumber As Long, sErrMsg As String)
If lErrNumber = 0 Then
    
    frmDisplay.AddToProcessList "Finished Delivery Without Errors"
Else
    
    frmDisplay.AddToProcessList "Finished Delivery With Error#" & lErrNumber & " " & sErrMsg
End If

DoEvents

GlobalTimer.Enabled = True
lblTimerState.Caption = ""
End Sub

Private Sub comAbout_Click()
frmAbout.Show
End Sub

Private Sub comAlarm_Click()
  SysAlarm.Show
End Sub

Private Sub comExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
quitting = False
ReadINIParams
TimeWait.Interval = gEmailTimeout * 1000  ' in ini FILE it is in seconds
GlobalTimer.Interval = QueueCheckInterval
LogMessage sysMessage, "Queue check time interval: " & QueueCheckInterval & " ms."
comAlarm.Enabled = gAlarmOn

ReDim pSavedActionReminder(ManyQueCount - 1)


End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim s As String
Dim i As Integer
Dim l As Long
Dim frmA As Form
Select Case UnloadMode
  Case vbFormControlMenu
    s = "form """ & Me.Caption & """ is closing from control menu."
  Case vbFormCode
    s = "form """ & Me.Caption & """ is closing from code."
  Case vbAppWindows
    s = "Windows session is closing."
  Case vbAppTaskManager
    s = "application is closing by Task Manager."
  Case vbFormOwner
    s = "form """ & Me.Caption & """ owner is closing."
  Case Else
    s = "unknown."
End Select
LogMessage sysMessage, App.EXEName & " is closing. Reason: " & s
If MsgBox("Do you want to close App?", vbDefaultButton2 + vbOKCancel) = vbOK Then
    Set frmA = New frmWait
    Load frmA
    frmA.Move Me.Left + (Me.Width - frmA.Width) / 2, Me.Top + (Me.Height - frmA.Height) / 2
    l = SetWindowPos(frmA.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE)
    DoEvents
    quitting = True
    For i = 0 To ManyQueCount - 1
      DoEvents
      Call RenameQue(ManyQTypes(i))
      Call CloseReadQue(ManyQTypes(i))
    Next i
    quitting = False
    Set EmailSender = Nothing
    If gAlarmOn Then
        Unload SysAlarm
    End If
    XnCloseDataBase
    CleanUpMutex
    LogMessage sysMessage, App.EXEName & " closed."
    Unload frmA
    End
Else
    Cancel = 1
    LogMessage sysMessage, "Closing application " & App.EXEName & " cancelled."
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Do you want to close App?", vbDefaultButton2 + vbOKCancel) = vbOK Then
    Set EmailSender = Nothing
    Dim i As Integer
    For i = 0 To ManyQueCount - 1
        Call CloseReadQue(ManyQTypes(i))
    Next
    If gAlarmOn Then
        Unload SysAlarm
    End If
    XnCloseDataBase
    CleanUpMutex
    LogMessage sysMessage, App.EXEName & " closed."
    End
Else
    Cancel = 1
    LogMessage sysMessage, "Closing application " & App.EXEName & " cancelled."
End If
End Sub

Private Sub GlobalTimer_Timer()
Dim iPos As Integer
Dim strProfileID As String
Dim strEmailAddr As String
Dim strMsgToEmail As String
Dim pagedMessage As String
Dim iQCounter As Integer
Static bSaveActRem As Boolean
Static isNotFirstTime As Boolean
Static iTic As Integer

On Error GoTo OOPS
If isNotFirstTime = False Then
    Errorform.Hide
    isNotFirstTime = True
End If
GlobalTimer.Enabled = False
lblTimerState.Caption = "Timer is Idle"
If iTic = 50 Then
    iTic = 0
Else
    iTic = iTic + 10
End If
For iQCounter = 0 To ManyQueCount - 1 ' -----
'lblWhatsUp.Caption = "Checking Queue Info " & QType.filename & String(iTic, "_")
    lblWhatsUp.Caption = "Checking Queue Info " & ManyQTypes(iQCounter).FileName & String(iTic, "_")
10  Do While Not CheckReadQue(ManyQTypes(iQCounter)) = False
        pagedMessage = ""
15      Call ReadQueLong(ManyQTypes(iQCounter), pagedMessage)
20      If UpdateReadQue(ManyQTypes(iQCounter)) = True Then
25          If ManyQTypes(iQCounter).nextWrite - ManyQTypes(iQCounter).nextRead > Q_LIMIT And Q_LIMIT > 0 Then
                ' Setup Action Reminder that the Q level is reaches LIMIT !!!
                If pSavedActionReminder(iQCounter) = False Then SaveActionReminder ManyQTypes(iQCounter)
                pSavedActionReminder(iQCounter) = True
            Else
                pSavedActionReminder(iQCounter) = False
            End If
            lblWhatsUp.Caption = "Retrieving Info from Queue ________________________"
30          iPos = InStr(QueRecord.PInfo, "+")
            If iPos <> 0 Then
                LogMessage sysMessage, "Q: nextWrite=" & ManyQTypes(iQCounter).nextWrite & " nextRead=" & ManyQTypes(iQCounter).nextRead & " QW_LIMIT=" & Q_LIMIT & " Pos=" & iPos & " Q#=" & iQCounter + 1 & " from " & ManyQueCount
                strProfileID = QueRecord.PExtension
35              strEmailAddr = Mid(QueRecord.PInfo, 1, iPos - 1)
                strMsgToEmail = GetRidOfJunkChar(pagedMessage)
                AddToProcessList "Q: " & ManyQTypes(iQCounter).queType & ". Email for profile: " & strProfileID & ", to: " & Trim$(GetRidOfJunkChar(strEmailAddr)) & ", msg: " & Trim$(GetRidOfJunkChar(strMsgToEmail)) & "..........."
                LogMessage sysMessage, "Q: " & ManyQTypes(iQCounter).queType & ". Email for profile: " & strProfileID & ", to: " & Trim$(GetRidOfJunkChar(strEmailAddr)) & ", msg: " & Trim$(GetRidOfJunkChar(strMsgToEmail))
40              Emailing strProfileID, strEmailAddr, strMsgToEmail
                
                TimeWait.Enabled = True
            
                'Exit Sub
                GoTo NextQ
            End If
        End If
DoEvents

NextQ:
    Loop
  ' rename queue:

50  Call RenameQue(ManyQTypes(iQCounter))

60  Call CloseReadQue(ManyQTypes(iQCounter))   ''' added 03/28/2011

Next '  For Loop ----------------------

ExitHere:
GlobalTimer.Enabled = True
lblTimerState.Caption = ""
Exit Sub
OOPS:
LogMessage ErrorMessage, MakeMessage(Err.Number, "Error line=" & Erl & ". " & Err.Description, "GlobalTimer", Me.Name)
Resume ExitHere
End Sub


Private Sub seeErrors_Click()

Errorform.Show
End Sub


Private Sub TimeWait_Timer()
'LogMessage sysMessage, "Waiting for <Notify> event"
TimeWait.Enabled = False
If GlobalTimer.Enabled = False Then
    ' Timer was idle....
    LogMessage sysMessage, "Timer was idle. Start the timer."
    GlobalTimer.Enabled = True
    lblTimerState.Caption = ""
End If
End Sub


