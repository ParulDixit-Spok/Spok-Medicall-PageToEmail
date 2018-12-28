Attribute VB_Name = "AppCommon"
Option Explicit

Public gIniPathFile As String

'Public DB As ADODB.Connection

Public sysMessage As MessageType

Public ErrorMessage As MessageType
Public pth  As String
Public Const MYDATEFORMAT = "mm/dd/yy hh:nn:ss"
'Public QType As ReadQueType

Public ManyQTypes() As ReadQueType   ' 10/29/12 ----------------
 
Public gEmailTimeout As Integer   ' in seconds

Public gEmailMode As String  ' SMTP   or OUTLOOK   or MAPI
Public gSMTPEmailFrom As String
Private gSubject As String
Private gUI As String
Private gPassword As String

Public QueueCheckInterval As Integer
Public EmailSender As Object  ' EMailSrv.OLEMailing
Public Q_LIMIT As Integer
Public gQ_id As String
Public XtendProfile As String
Public DbName As String
Public DBPassword As String
Public DBUserName As String
Public gAlarmOn As Boolean  ' XTEND ALARM
Public CustomMsgID As Integer  ' Message for Alarm to be sent in case of locking in rename routine.
Public gMinutesBeforeMidnightToInitializeQueue As Long
Public quitting As Boolean

Public ManyQueCount As Integer  '' 10/29/12 Multiple Q count
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Function GetUniqueNumber()
Dim lgNumber As Long
Dim fmt As String
' Added 12/03/2014 to add RecordNumber to a colllection for Emailing so it can process many emails at once, like from delivery center
' before it was 0 all the time
On Error Resume Next

Randomize
lgNumber = (10 ^ 3) - 1

lgNumber = Val((lgNumber * Rnd) + 1)
Dim tmrMilisec As Long
tmrMilisec = Mid(CStr(Timer), InStr(CStr(Timer), ".") + 1)

GetUniqueNumber = lgNumber + tmrMilisec


End Function

Public Sub SetAllTables()

    OpenXKMTable (gIniPathFile)
    SetPagersTable (gIniPathFile)
    SetMsgTable (gIniPathFile)
    SetCountTable (gIniPathFile)


End Sub


Public Sub Emailing(parProfile As String, parEmail As String, parMsg As String)

Dim res As Integer
On Error GoTo OOPS

100 Set frmDisplay.CallBack = New clsCallBack
110 frmDisplay.CallBack.ProfileID = parProfile
115 EmailSender.RecordNumber = GetUniqueNumber() '' some long unique number ( a must otherwise it will not add new member to collection to be processed by timer in Emailsrv)
120 EmailSender.EMailAddr = parEmail
130 EmailSender.JobType = "G"
140 EmailSender.MessageToDeliver = parMsg
150 EmailSender.PassWord = gPassword
160 EmailSender.username = gUI
170 EmailSender.Subject = gSubject
172 EmailSender.Mode = gEmailMode
175 EmailSender.SMTP_FromEmail = gSMTPEmailFrom
180 res = EmailSender.SendEMail(frmDisplay.CallBack)

190 frmDisplay.AddToProcessList "Info Sent to email server for profileID = " & parProfile
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, MakeMessage(Err.Number, Err.Description, "Emailing", "AppCommon", Erl)
Resume ExitHere
End Sub



Sub Main()
'if App.PrevInstance Then End
If IsPrevInstanceRunning(App.EXEName) Then End

gOperator = "PgToEmail"

Set EmailSender = CreateObject("EMailSrv.OLEMailing")      'New EMailSrv.OLEMailing
frmDisplay.Show
Errorform.Show

End Sub





Public Sub ReadINIParams()
Dim arrQueIDs() As String ' 10/28/12  serve many Qs
Dim i As Integer
Dim msg As String, FileName As String

On Error GoTo OOPS

pth = App.Path
If Right(pth, 1) <> "\" Then pth = pth & "\"

gIniPathFile = pth & App.EXEName & ".ini"

InitLog sysMessage, gIniPathFile, App.EXEName & ".log", "", 100000
InitLog ErrorMessage, gIniPathFile, App.EXEName & ".err", ""
'LogMessage sysMessage, "Testing"
FileName = App.Path & "\" & App.EXEName & ".exe"
msg = "Starting " & FileName & " " & App.Major & "." & App.Minor & ".0." & App.Revision & " from " & Format(FileDateTime(FileName), "MM/DD/YYYY hh:nn:ss") & " on " & MyComputerName
LogMessage sysMessage, msg

ErrorMessage.displayMsgBox = LCase(Command$) = "debug"
gQ_id = GetIniString("QUEUE", "TYPE", "%10", gIniPathFile)
arrQueIDs = Split(gQ_id, "_")

ManyQueCount = UBound(arrQueIDs) + 1

ReDim Preserve ManyQTypes(ManyQueCount - 1)

For i = 0 To ManyQueCount - 1
    Dim strQ_Id As String: strQ_Id = arrQueIDs(i)
    If InitReadQue(gIniPathFile, strQ_Id, ManyQTypes(i)) = 0 Then
        LogMessage ErrorMessage, "Cannot Open Queue File " & gQ_id & " in ReadINIParams"
        Errorform.Show
    End If
Next

'If InitReadQue(gIniPathFile, gQ_id, QType) = 0 Then
'    LogMessage ErrorMessage, "Cannot Open Queue File " & gQ_id & " in ReadINIParams"
'    Errorform.Show
'End If
    
QueueCheckInterval = GetIniVal("QUEUE", "READING_INTERVAL", 1000, gIniPathFile)

gSubject = GetIniString("EMAIL", "SUBJECT", "", gIniPathFile)
gUI = GetIniString("EMAIL", "USERNAME", "", gIniPathFile)
gPassword = GetIniString("EMAIL", "PASSWORD", "", gIniPathFile)
gEmailTimeout = GetIniVal("EMAIL", "EMAIL_TIMEOUT", 10, gIniPathFile)
gEmailMode = GetIniString("EMAIL", "MODE", "OUTLOOK", gIniPathFile, True)

gSMTPEmailFrom = GetIniString("EMAIL", "EmailFrom", "", gIniPathFile, True)

Q_LIMIT = GetIniVal("QUEUE", "WARNING_LIMIT", 0, gIniPathFile)
gMinutesBeforeMidnightToInitializeQueue = CLng(GetIniString("QUEUE", _
                                    "MinutesBeforeMidnightToInitializeQueue", _
                                    "5", _
                                    gIniPathFile))

XtendProfile = GetIniString("XN", "XTEND_PROFILE", "XTEND", gIniPathFile, True)
DbName = GetIniString("XN", "DATABASE", "Smart Answer", gIniPathFile, True)
DBPassword = GetIniString("XN", "DB_PASSWORD", "", gIniPathFile, True)
DBUserName = GetIniString("XN", "DB_USER", "", gIniPathFile, True)

SetXtendAlarm

LogMessage sysMessage, "E-mail userID = " & gUI & ", Password = " & gPassword
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, MakeMessage(Err.Number, Err.Description, "ReadINIParams", "AppCommon")
Resume ExitHere
End Sub



Public Function MakeMessage(parErr As Long, parError As String, _
        parRoutine As String, parModule As String, Optional eLine As Integer) As String
        
Dim X As String

On Error Resume Next
If eLine <> 0 Then
    X = "Err: " & CStr(parErr) & " - " & parError & " in " & parModule & "." & parRoutine & "(), line# " & CStr(eLine)
Else
    X = "Err: " & CStr(parErr) & " - " & parError & " in " & parModule & "." & parRoutine & "()"
End If
MakeMessage = X


End Function

Public Sub SaveActionReminder(thisQType As ReadQueType)
Dim iTry As Integer
Dim lgLastMsgNumber As Long
Dim strMessage As String
Dim strInformation As String
Const ALARM_TYPE = "1 "
Const RECORD_TYPE = 6
Dim strProfileID As String * 10
On Error GoTo OOPS

RSet strProfileID = XtendProfile
strMessage = "Queue level is reached for Que File: " & thisQType.FileName
If DB.State = adStateClosed Then
    XnOpenDataBase DbName, DBUserName, DBPassword
    SetAllTables
End If

lgLastMsgNumber = SaveNewMsg(strProfileID, "M", _
    gOperator, strMessage, True, _
    "PageToEmail Utility")

If lgLastMsgNumber > 0 Then
    strInformation = ALARM_TYPE & Space(10 - Len(Trim(strProfileID))) & Trim(strProfileID) & CStr(lgLastMsgNumber)
    
    CatchNewRecord RECORD_TYPE, strInformation, Format(Now, "yyyymmddhhnn"), strProfileID
    LogMessage sysMessage, "Saved Action Reminder for Profile: " & XtendProfile & " and MessageNumber: " & CStr(lgLastMsgNumber)
    Errorform.Show
End If
ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, MakeMessage(Err.Number, Err.Description, "AppCommon", "SaveActionReminder")
Resume ExitHere
End Sub


Private Sub SetXtendAlarm()
Dim iHowOftenToPutInSched As Integer
Dim iAlarmTimeSchedule As Integer
Dim strAlarmMsg1 As String
Dim strAlarmMsg2 As String
Dim iRepeatInterval As Integer

gAlarmOn = GetIniBool("XTEND ALARM", "AlarmOn", "FALSE", gIniPathFile)
iHowOftenToPutInSched = GetIniVal("XTEND ALARM", "HowOftenPutInSchedule", 180, gIniPathFile)
iAlarmTimeSchedule = GetIniVal("XTEND ALARM", "AlarmTimerSchedule", 80, gIniPathFile)
strAlarmMsg1 = GetIniString("XTEND ALARM", "ABNORMAL MSG ID", "", gIniPathFile)
strAlarmMsg2 = GetIniString("XTEND ALARM", "CLOSE MSG ID", "", gIniPathFile)
CustomMsgID = GetIniString("XTEND ALARM", "WARNING MSG ID", 0, gIniPathFile)

iRepeatInterval = GetIniVal("XTEND ALARM", "RepeatInterval", 23, gIniPathFile)

SysAlarm.SetObjStatPara iHowOftenToPutInSched, iAlarmTimeSchedule, strAlarmMsg1, strAlarmMsg2, iRepeatInterval, _
    DbName, DBUserName, DBPassword

If gAlarmOn Then
    Load SysAlarm
Else
    SysAlarm.Timer1.Enabled = False
    SysAlarm.DelPreMsg
End If

End Sub

Public Function MyComputerName() As String
Dim L As Long
Dim i As Integer
Dim cn As String * 200, s As String
On Error Resume Next
MyComputerName = vbNullString
L = GetComputerName(cn, 200)
If L > 0 Then
  i = 1
  s = Left$(cn, 1)
  Do While Asc(s) > Asc(Space(1))
    MyComputerName = MyComputerName & s
    i = i + 1
    s = Mid$(cn, i, 1)
  Loop
End If
MyComputerName = Trim$(MyComputerName)
End Function

