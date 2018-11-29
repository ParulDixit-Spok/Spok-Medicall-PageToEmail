Attribute VB_Name = "Read_Page_Q"
Option Explicit
'
'
'   XoRQue.bas          Routines to Read and work with paging queues
'
'   USAGE:              Setup:  Define a structure of type ReadQueType for Every Que to be open
'                               Initialize XnDiags
'                               Call InitReadQue for every ReadQueType
'
'                                   Private q42Info As ReadQueType
'
'                                   Call InitLog(errorMessage, pth & "test.INI", "test.ERR", "Error")
'                                   Call InitLog(sysMessage, pth & "test.INI", "test.MSG", "System")
'                                   Call InitReadQue(pth & "test.INI", "%42", q42Info)
'
'                       Process: Call CheckReadQue, UpdateReadQue, ReadQue, and RenameQue
'
'                                    If CheckReadQue(q42Info) = True Then
'                                        Label1.Caption = "Read: " & q42Info.nextRead
'                                        Label2.Caption = "Write: " & q42Info.nextWrite
'                                        If UpdateReadQue(q42Info) = True Then
'                                            Call ReadQue(q42Info)
'
'                                            !!! YOUR CODE HERE !!!
'
'                                        End If
'                                    Else
'                                        Call RenameQue(q42Info)
'                                    End If

'
'                       Cleanup: Call CloseReadQue for Every ReadQueType
'
'                                    Call CloseReadQue(q42Info)
'                                    CloseLog sysMessage
'                                    CloseLog errorMessage
'                                    Unload Errorform
'
'   Ini Settings:       [XN]
'                       ' Path for Que Files
'                       QUE_PATH=N:\NOTIS
'                       ' Max Entries in queue for rename
'                       QUE_LIMIT = 10
'                       '   TRUE delete queue files instead of renaming
'                       QUE_REMOVE = FALSE
'
'   Requered Moduals:   Standard Diagnostics and Ini: Xnini.bas, Xndiags.frm, Xndiags.bas, CSUB.bas, Generict.bas
'
'   Created:    10/13/98    Joseph Slawinski
'

'
'   ReadQueType: Contains processing info for every queue to be processed
'
Type ReadQueType
    fileNumber As Integer       '   Number for open file
    nextRead As Long         '   Next record to read  (Changed from Int to Long to hold QUE_LIMIT > 100 , because 100*226(Len9QueRecord) > 32 000
    nextWrite As Long        '   Next record to write (Changed from Int to Long to hold QUE_LIMIT > 100 , because 100*226(Len9QueRecord) > 32 000
    renameCount As Integer      '   Tracks rename file names
    FileName As String          '   Queue file name i.e. N:\XTEND\NOPG42.QUE
    queType As String           '   Queue identifier i.e. %42
End Type

'
'   PageQueHeader:  Strores page que header information, first 20 bytes
'
Type PageQueHeader
    GetPointer As Integer       '   Next Read
    Temp1 As Integer
    PError1 As Integer
    PError2 As Integer
    PType As String * 2
    PutPointer As Integer       '   Next Write
    temp2 As Integer
    Junk As String * 6
End Type
Dim QueHeader As PageQueHeader

'
'   PageQueRecord:  Stores Queue Record Information
'
Type PageQueRecord
    PType As String * 2         '   Page Type i.e. 41 for %41
    PStatus As String * 2
    PDatein As String * 10      '   date of page record
    PTimein As String * 8       '   time of page record
    PDateout As String * 10
    PTimeout As String * 8
    PExtension As String * 8    '   profile id
    PExtid As String * 7        '   pager id
    PIdin As String * 10        '   user id
    Packed As String * 1
    Packtime As Integer
    PPointer As Integer
    PPrinted As String * 1
    PVoice As String * 5
    PVoicef As Integer
    PInfo As String * 148       '   page information
End Type
Global QueRecord As PageQueRecord
Public QueueOpened As Boolean
Type PageQueBtrRecord
    interfaceType As String * 10    '   Used to identify interface used to send page
    priority As String * 2          '   0 highest, 99 lowest
    initiatedDate As String * 10
    initiatedTime As String * 8
    initiatedId As String * 10
    ProfileID As String * 10
    PagerId As String * 10
    PageType As String * 10         '   Identifies type of pager
    voiceFlag As String * 1         '   Y indicates voice page
    voiceFormat As String * 15      '   Encoding format such as wav, dialogic, ect. Blank idicates default
    FileName As String * 120        '   filename of text to page or voice file to page
    pageInfo As String * 300        '   page information + message. As used in queue file inteface
    reserved As String * 194        '   reserved for future expansion
End Type
Dim QueBtrRecord As PageQueBtrRecord
    
Private XnPageQTable As New ADODB.Recordset
    
Private pageQueueBtrvTable As String    ' Table name for Page queue
Private QuePath As String       '   path for .que files
Private querecordlen As Integer, pointerreclen  As Integer      ' lenght of querecord and header
Private waittime As Variant
Private queLimit As Integer     '   max entries in queue for rename
Private queRemove As Integer    '   True = delete queues instead of renaming

Global Const lockwait = 15

Function InitReadQue(iniFileName As String, qId As String, qInfo As ReadQueType) As Integer
'
'   Opens and initializes queue file.
'
'   Parameters: iniFileName:    path and name of inifile
'               qId:            queue id i.e. %42
'               qInfo           ReadQueType Structure
'
'   Returns:    True    success
'               False   error
'
On Error GoTo initqueerror

InitReadQue = False
querecordlen = Len(QueRecord)
pointerreclen = Len(QueHeader)

QuePath = GetIniString("XN", "QUE_PATH", ".", iniFileName)
queLimit = GetIniVal("XN", "QUE_LIMIT", 150, iniFileName)
queRemove = GetIniBool("XN", "QUE_REMOVE", "FALSE", iniFileName)
pageQueueBtrvTable = GetIniString("XN", "PAGE_QUEUE_BTRV_TABLE", "", iniFileName)

qInfo.queType = qId
qInfo.renameCount = 0
qInfo.FileName = ""


    waittime = DateAdd("s", lockwait, Now)
    
    qInfo.queType = qId
    qInfo.renameCount = 0
    qInfo.FileName = QuePath & "\NOPG" & Right$(qId, 2) & ".QUE"
    
    qInfo.fileNumber = FreeFile
100     Open qInfo.FileName For Binary Shared As qInfo.fileNumber
    
    If LOF(qInfo.fileNumber) < 20 Then
        QueHeader.GetPointer = 0
        QueHeader.PError1 = 0
        QueHeader.PError2 = 0
        QueHeader.PType = "1 "
    '   QueHeader.putpointer = 1
        QueHeader.PutPointer = 0
        QueHeader.Junk = String$(Len(QueHeader.Junk), " ")
        Put qInfo.fileNumber, 1, QueHeader
    End If
    
    InitReadQue = True
    
    LogMessage sysMessage, "Que open: " & qInfo.FileName & " #: " & qInfo.fileNumber


1000: Exit Function

initqueerror:

If Err = 70 Then
    Do While waittime > Now
        'DoEvents
        Resume
    Loop
    LogMessage ErrorMessage, "Lock fail InitReadQue " & Erl & " Que file locked"
    Resume 1000
End If

    LogMessage ErrorMessage, "InitReadQue error =" & Error & " ERL =" & Erl

    Resume 1000
End Function

Sub CloseReadQue(qInfo As ReadQueType)
'
'   Closes Queue File
'
'   Parameters: qInfo           ReadQueType Structure
'
On Error GoTo closereadqueerror

If pageQueueBtrvTable = "" Then
    Close #qInfo.fileNumber
100:
    qInfo.fileNumber = 0
    If QueueOpened Then
      LogMessage sysMessage, "Que Closed: " & qInfo.FileName
      QueueOpened = False
    End If
End If
exitclosereadque:

    Exit Sub

closereadqueerror:

    LogMessage ErrorMessage, "CloseReadQue error =" & Error & " ERL =" & Erl

    Resume exitclosereadque

End Sub

Function CheckReadQue(qInfo As ReadQueType) As Single
'
'   Checks for unprocesed queue records
'
'   Parameters: qInfo           ReadQueType Structure
'
'   Returns:    True    Unprocessed record in queue
'               False   queue empty
'-------------------------------------------------------------------------
'---------- DO NOT CHANGE LINE NUMBERS !!!    ---------------------------
'--------------------------------------------------------------------------
    On Error GoTo checkreadqueerror
10  QueueOpened = False
100 If pageQueueBtrvTable = "" Then
110    If qInfo.fileNumber > 0 Then
120        'check to see if que needs service
130        Lock qInfo.fileNumber, 11 To 20 ' do not change this line number - 130 !!!
250        Get qInfo.fileNumber, 1, QueHeader 'do not change this line number -250 !!!
260        Unlock qInfo.fileNumber, 11 To 20
        
270        qInfo.nextRead = QueHeader.GetPointer ' if wanttto point to a certain rec put QueHeader.GetPointer some val < then QueHeader.PutPointer
280        qInfo.nextWrite = QueHeader.PutPointer
290        If QueHeader.PutPointer > QueHeader.GetPointer Then
300            CheckReadQue = True
               QueueOpened = True
305            LogMessage sysMessage, "CheckReadQue: Q " & qInfo.FileName & " nextWrite=" & qInfo.nextWrite & " nextRead=" & qInfo.nextRead & " fileNumber=" & qInfo.fileNumber
310        Else
320            CheckReadQue = False
330        End If
335    Else
340
350        qInfo.fileNumber = FreeFile
'355        LogMessage sysMessage, "CheckReadQue: Q " & qInfo.FileName & " # " & qInfo.fileNumber
360        Open qInfo.FileName For Binary Shared As qInfo.fileNumber
370        If LOF(qInfo.fileNumber) < 20 Then
380            QueHeader.GetPointer = 0
390            QueHeader.PError1 = 0
400            QueHeader.PError2 = 0
410            QueHeader.PType = "1 "
        '        QueHeader.putpointer = 1
420            QueHeader.PutPointer = 0
430            QueHeader.Junk = String$(Len(QueHeader.Junk), " ")
440            Put qInfo.fileNumber, 1, QueHeader
442            LogMessage sysMessage, "CheckReadQue: Header is initialized!"
445        Else
446            Lock qInfo.fileNumber, 11 To 20 ' do not change this line number - 130 !!!
447            Get qInfo.fileNumber, 1, QueHeader 'do not change this line number -250 !!!
448            Unlock qInfo.fileNumber, 11 To 20
449            qInfo.nextRead = QueHeader.GetPointer ' if wanttto point to a certain rec put QueHeader.GetPointer some val < then QueHeader.PutPointer
450            qInfo.nextWrite = QueHeader.PutPointer
452            If QueHeader.PutPointer > QueHeader.GetPointer Then
453                  CheckReadQue = True
                     QueueOpened = True
454                  LogMessage sysMessage, "CheckReadQue: Q " & qInfo.FileName & " nextWrite=" & qInfo.nextWrite & " nextRead=" & qInfo.nextRead
455            Else
458                  CheckReadQue = False
459            End If
460        End If
480    End If
470 Else
    
 End If

exitcheckreadque:

    Exit Function

checkreadqueerror:
    Dim strErr As String
    Dim iErr As Long
    Dim lineErr As Integer
    strErr = Err.Description: iErr = Err.Number: lineErr = Erl
    If lineErr = 250 Or lineErr = 130 Then ' problem with GET statement after Lock, so unlock , or can't lock because it was locked before (TK 09/28/06)
        On Error Resume Next
         LogMessage ErrorMessage, "Cannot lock Queue file."
         'Unlock qInfo.fileNumber, 11 To 20  - no need to unlock - it will give an error
         LogMessage ErrorMessage, "Trying to Re-Open the queue file...." & qInfo.FileName
         Call CloseReadQue(qInfo)
         LogMessage ErrorMessage, "File is closed."
         If InitReadQue(gIniPathFile, gQ_id, qInfo) = 0 Then
            LogMessage ErrorMessage, "Cannot Open Queue File " & gQ_id & " in CheckReadQue"
         Else
           LogMessage ErrorMessage, "Queue file is open now. Exiting CheckReadQue() without reading info."
           CheckReadQue = False
           Resume exitcheckreadque
         End If
    End If
    LogMessage ErrorMessage, "CheckReadQue error =" & strErr & ", line# " & lineErr
    CheckReadQue = False
    Resume exitcheckreadque

End Function

Function ReadQue(qInfo As ReadQueType) As Boolean
'
'   Reads next queue record
'
'   Parameters: qInfo           ReadQueType Structure
'
Dim getpagepointer As Long, pluspointer As Integer
Dim strTemp As String

On Error GoTo readqueerror
    
105    ReadQue = False
    
100    getpagepointer = pointerreclen + 1 + (qInfo.nextRead * querecordlen)
120    Get qInfo.fileNumber, getpagepointer, QueRecord
130    ReadQue = True
   
    
exitreadque:

    Exit Function

readqueerror:

    LogMessage ErrorMessage, "ReadQue error =" & Error & ", Err line: " & Erl
    Resume exitreadque
    
End Function

Public Function ReadQueLong(ByRef qInfo As ReadQueType, ByRef strPagedMessage As String)
'  Reads next queue record that have bigger length than default combining several records in 1
'  based on a continuation flag #
'
'  Parameters: qInfo           ReadQueType Structure
'
Dim getpagepointer As Long, pluspointer As Integer
Dim strTemp As String
Dim strLongMessage As String
Dim iPlusPos As Integer
On Error GoTo readqueerror
    
90    ReadQueLong = False
100    Do While True
105        getpagepointer = pointerreclen + 1 + (qInfo.nextRead * querecordlen)
110        Get qInfo.fileNumber, getpagepointer, QueRecord
120        LogMessage sysMessage, "Retrieve from Queue: " & qInfo.FileName & " -> " & Trim$(GetRidOfJunkChar(QueRecord.PInfo))
130        iPlusPos = InStr(QueRecord.PInfo, "+")
140        If iPlusPos > 0 Then
150            strLongMessage = strLongMessage & Mid(QueRecord.PInfo, iPlusPos + 1)
160        End If
170        If QueRecord.Packed <> "#" Then ' !!!
            
180            Exit Do
190        Else  ' for #
200            qInfo.nextRead = qInfo.nextRead + 1
205            LogMessage sysMessage, "ReadQueLong: new nextRead=" & qInfo.nextRead
210        End If
220    Loop
    
230    strPagedMessage = strLongMessage
    
240    ReadQueLong = True
    
    
exitreadque:

    Exit Function

readqueerror:

    LogMessage ErrorMessage, "ReadQueLong error =" & Error & ", Err line=" & Erl
    LogMessage ErrorMessage, "Checking variable getpagepointer=" & getpagepointer
    LogMessage ErrorMessage, "Checking variable nextRead=" & qInfo.nextRead
    LogMessage ErrorMessage, "Checking variable querecordlen=" & querecordlen
    LogMessage ErrorMessage, "Checking variable pointerreclen=" & pointerreclen
    Resume exitreadque
    Resume
End Function


Function UpdateReadQue(qInfo As ReadQueType) As Single
'
'   Increments queue nextread pointer
'
'   Parameters: qInfo ReadQueType Structure
'
'   Returns:    True    Successful
'               False   Error
'
On Error GoTo advanceerror
    If pageQueueBtrvTable <> "" Then
        UpdateReadQue = True
        Exit Function
    End If

    UpdateReadQue = False
    waittime = DateAdd("s", lockwait, Now)

1100: Lock #qInfo.fileNumber, 11 To 20
1200: Get qInfo.fileNumber, 1, QueHeader

        'QueHeader.GetPointer = QueHeader.GetPointer + 1
        QueHeader.GetPointer = qInfo.nextRead + 1

    If QueHeader.GetPointer <= QueHeader.PutPointer Then
1250:   Put #qInfo.fileNumber, 1, QueHeader.GetPointer
        UpdateReadQue = True
        LogMessage sysMessage, "UpdateReadQue: new nextRead=" & QueHeader.GetPointer
    Else
        LogMessage sysMessage, "Update Q " & qInfo.FileName & " Pointer Error GET (" & QueHeader.GetPointer & ") > PUT (" & QueHeader.PutPointer & ")"
        LogMessage ErrorMessage, "Update Q " & qInfo.FileName & " Pointer Error GET (" & QueHeader.GetPointer & ") > PUT (" & QueHeader.PutPointer & ")"
    End If
    
1300: Unlock #qInfo.fileNumber, 11 To 20

2000: Exit Function

advanceerror:

    If Err = 70 Then
        Do While waittime > Now
    '        DoEvents
            Resume
        Loop
        LogMessage ErrorMessage, "UpdateReadQue: Lock fail " & Erl
        If Erl <> 1100 Then
            Unlock #qInfo.fileNumber, 11 To 20
        End If
        Resume 2000
    End If

    If Erl <> 1100 Then
        Unlock #qInfo.fileNumber, 11 To 20
    End If
    LogMessage ErrorMessage, "UpdateReadQue Error line=" & Erl & " Error: " & Error$
    Resume 2000

End Function

Sub RenameQue(qInfo As ReadQueType)
'
'   Renames or removes fully processed queue after queLimit records are processed
'
'   Parameters: qInfo           ReadQueType Structure
'
Dim pointer As Integer, updateflag As Integer, currentdate As Variant
Dim fileNum As Integer, newname As String
Dim intTry As Integer
Dim t As Date

On Error GoTo renamequeerror
    If pageQueueBtrvTable <> "" Then Exit Sub
    If Not quitting Then
      t = Format(Now, "mm/dd/yyyy") & " 23:59:59"
      If DateDiff("n", Now, t) < 0 Or DateDiff("n", Now, t) >= gMinutesBeforeMidnightToInitializeQueue Then
        If qInfo.nextRead < queLimit Then Exit Sub
      End If
    End If
    If FileLen(qInfo.FileName) < 200 Then
      Exit Sub
    End If
    currentdate = Date
10  If queRemove = False Then
'        LogMessage sysMessage, "Renaming Queue: " & qInfo.FileName
        'kill the target if it exits
20      Do While qInfo.renameCount < 999
            qInfo.renameCount = qInfo.renameCount + 1
            newname = QuePath & "\R" & Format$(Now, "MMDD") & Format$(qInfo.renameCount, "000") & ".Q" & Right$(qInfo.queType, 2)
25          If Dir$(newname) > "" Then
                If qInfo.renameCount = 999 Then
30                Kill newname
                  LogMessage sysMessage, "Killing old Queue: " & newname
                  If quitting Then
                    DoEvents
                  End If
                End If
            Else
                Exit Do
            End If
        Loop
    Else
        LogMessage sysMessage, "Deleting queue file: " & qInfo.FileName
    End If
    'check record pointers
    
40  Lock qInfo.fileNumber, 11 To 20
45  Get qInfo.fileNumber, 1, QueHeader
50  Unlock qInfo.fileNumber, 11 To 20
    If quitting Then
      DoEvents
    End If
60  If QueHeader.PutPointer = QueHeader.GetPointer Then
            
        waittime = DateAdd("s", lockwait, Now)

65      Get qInfo.fileNumber, 1, QueHeader
        If QueHeader.PutPointer = QueHeader.GetPointer Then
70          Close #qInfo.fileNumber
            qInfo.fileNumber = 0
100:
            If queRemove = False Then
                Name qInfo.FileName As newname
                LogMessage sysMessage, "Renaming queue file: " & qInfo.FileName & " to " & newname
                fileNum = FreeFile
200:            Open newname For Binary Shared As fileNum
                Get fileNum, 1, QueHeader
                If QueHeader.PutPointer > QueHeader.GetPointer Then
                    LogMessage ErrorMessage, "Page request missed in file " & newname
                End If
                Close fileNum
            Else
250             Kill qInfo.FileName
            End If
        End If
        If quitting Then
          DoEvents
        End If
        qInfo.fileNumber = FreeFile
300:    Open qInfo.FileName For Binary Shared As qInfo.fileNumber
        If LOF(qInfo.fileNumber) < 20 Then
            QueHeader.GetPointer = 0
            QueHeader.PError1 = 0
            QueHeader.PError2 = 0
            QueHeader.PType = "1 "
        '        QueHeader.putpointer = 1
            QueHeader.PutPointer = 0
            QueHeader.Junk = String$(Len(QueHeader.Junk), " ")
350         Put qInfo.fileNumber, 1, QueHeader
        End If
        qInfo.nextRead = 0
        qInfo.nextWrite = 0
400:    LogMessage sysMessage, "RenameQue: Header is initialized!"
    End If
    If queRemove = False Then
        LogMessage sysMessage, "Queue: " & qInfo.FileName & " renamed as: " & newname
    Else
        LogMessage sysMessage, "Queue: " & qInfo.FileName & " deleted"
    End If
1000: Exit Sub

renamequeerror:
    If Erl = 100 Or Erl = 200 Then
         If intTry < 4 Then
            LogMessage ErrorMessage, "Error in renaming/removing Queue, trying again... [" & CStr(intTry + 1) & "] times."
            intTry = intTry + 1
            Delay 100
            Resume
        End If
    End If
    
    If gAlarmOn Then
        SysAlarm.PostMsgNow (CustomMsgID)
    End If
    
    LogMessage ErrorMessage, "RenameQue: Error line= " & Erl & " Error: " & Error$
    Resume 1000
End Sub


Sub XnCloseRQueueBtrv()

On Error GoTo OOPS

If pageQueueBtrvTable <> "" Then
    XnPageQTable.Close
    LogMessage sysMessage, "Page Queue Btrv Table " & pageQueueBtrvTable & " Closed"
End If

ExitHere:
Exit Sub

OOPS:
LogMessage ErrorMessage, "Close " & pageQueueBtrvTable & " Error" & CStr(Err) & " in XnCloseRQueueBtrv " & Error$
Resume ExitHere

End Sub

