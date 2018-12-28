Attribute VB_Name = "ModScheduleADO"
Option Explicit
Public ScheduleTable As New ADODB.Recordset
Public Const NameOfSchedule = "Schedule"

Function CatchNewRecord(TheRecordType As Integer, TheInfo As String, TheHeader As String, TheExtension As String) As Boolean
        Dim sSQL As String
        Dim intCounter As Integer
        Dim MaxCounter As Integer
        Dim TheId As String * 10
        Dim j As Integer
        
        On Error GoTo OOPS
        LogMessage sysMessage, "CatchNewRecord Enter :" & Now
        RSet TheId = Trim(TheExtension)
StartAgain:
        If DB.State <> adStateOpen Then
           XnOpenDataBase DbName, DBUserName, DBPassword
           'OpenAllTables Parameter.pPath
           SetAllTables
        End If
    
        With ScheduleTable
             If .State = adStateOpen Then .Close
            .ActiveConnection = DB
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            sSQL = "Select count(*) as c1 from " & NameOfSchedule & " Where RecordType = " & CStr(TheRecordType)
            .Open sSQL
            intCounter = .Fields("c1")
            .Close
        End With
        ''Set ScheduleTable = Nothing
        If intCounter = 0 Then
           intCounter = 1
        Else
            With ScheduleTable
                If .State = adStateOpen Then .Close
               .ActiveConnection = DB
               .CursorType = adOpenDynamic
               .LockType = adLockOptimistic
               sSQL = "Select Max(RecordCount) as m1 from " & NameOfSchedule & " Where RecordType = " & CStr(TheRecordType)
               .Open sSQL
               MaxCounter = .Fields("m1")
               .Close
            End With
            If intCounter > MaxCounter Then
               intCounter = MaxCounter + 1
            Else
               intCounter = intCounter + 1
            End If
        
        End If
        With ScheduleTable
             If .State = adStateOpen Then .Close
             .Open NameOfSchedule, DB, adOpenDynamic, adLockOptimistic, adCmdTable
             .AddNew
             .Fields("Header") = TheHeader
             .Fields("RecordType") = TheRecordType
             .Fields("TimeType") = 0
             .Fields("generalinformation") = TheInfo
             .Fields("Frequency") = 0
             .Fields("ScheduleTime") = ""
             .Fields("Extension") = TheId
             .Fields("RecordCount") = intCounter
             .Update
             .Close
        End With
        Set ScheduleTable = Nothing
        CatchNewRecord = True
        LogMessage sysMessage, "CatchNewRecord Exit :" & Now
exiterror:
   Exit Function
OOPS:
        LogMessage ErrorMessage, "CatchNewRecord :" & Err.Description
        j = j + 1
        If j <= 2 Then
        If DB.Errors.Count <> 0 Then
           If DB.Errors(0).NativeError <> 0 Then
              If ScheduleTable.State = adStateOpen Then
                 ScheduleTable.Close
                 Set ScheduleTable = Nothing
              End If
              CloseAllTables
              If DB.State = adStateOpen Then
                 DB.Close
                 Set DB = Nothing
              End If
              Delay 500
              
              Resume StartAgain
           End If
        End If
       End If
       Resume exiterror
End Function

Public Function GetRecordByKey(TheRecordType As Integer, TheKeyNumber As Integer, ComparedBy As String, KeyValue As Variant, Optional KeyValue2) As ColSchedule
        Dim sSQL As String
        Dim Header1 As String
        Dim RecordType1 As Integer
        Dim Extension1 As String
        Dim RecordCount1 As Integer
        Dim GeneralInformation1 As String
        Dim StrParameters As String
        Dim j As Integer
        Dim k As Integer
        On Error GoTo OOPS
        Set GetRecordByKey = New ColSchedule
StartAgain:
        If DB.State <> adStateOpen Then
            XnOpenDataBase DbName, DBUserName, DBPassword
            'OpenAllTables Parameter.pPath
            SetAllTables
        End If
        With ScheduleTable
             If .State = adStateOpen Then .Close
            .ActiveConnection = DB
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
             If IsMissing(KeyValue2) = True Then
                sSQL = "select * from " & NameOfSchedule & " Where RecordType " & ComparedBy & " '" & KeyValue & "' " & " Order by Header "
             Else
                 sSQL = "select * from " & NameOfSchedule & " Where RecordType " & ComparedBy & " " & KeyValue & " " & _
                            " And Extension " & ComparedBy & " '" & KeyValue2 & "' " & " Order by Header "
                
             End If
             .Open sSQL
             If .EOF = True And .BOF = True Then
                'no records in the schedule table
             Else
                Do
                   Header1 = .Fields("Header")
                   RecordType1 = .Fields("RecordType")
                   Extension1 = .Fields("Extension")
                   RecordCount1 = .Fields("RecordCount")
                   GeneralInformation1 = .Fields("GeneralInformation")
                   GetRecordByKey.Add Header1, RecordType1, Extension1, RecordCount1, GeneralInformation1
                   .MoveNext
                
                Loop While .EOF <> True
                .Close
             End If
        End With
        Set ScheduleTable = Nothing
exiterror:
        Exit Function
OOPS:
        LogMessage ErrorMessage, "GetRecorByKey: " & Err.Description
        k = k + 1
        If k <= 2 Then
        If DB.Errors.Count <> 0 Then
           If DB.Errors(0).NativeError <> 0 Then
              If ScheduleTable.State = adStateOpen Then
                 ScheduleTable.Close
                 Set ScheduleTable = Nothing
              End If
              CloseAllTables
              If DB.State = adStateOpen Then
                 DB.Close
                 Set DB = Nothing
              End If
              Delay 500
              
              Resume StartAgain
           End If
        End If
       End If
       Resume exiterror
        
End Function

Sub DeleteScheduleRecord(intRecordType As Integer, intRecCount As Integer, strExtension As String)
    Dim sSQL As String
    Dim temExten As String * 10
    Dim j As Integer
    On Error GoTo OOPS
     
     RSet temExten = Trim(strExtension)
StartAgain:
     If DB.State <> adStateOpen Then
           XnOpenDataBase DbName, DBUserName, DBPassword
           'OpenAllTables Parameter.pPath
           SetAllTables
     End If
     With ScheduleTable
             If .State = adStateOpen Then .Close
            .ActiveConnection = DB
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            sSQL = "Select * from " & NameOfSchedule & " Where RecordType =" & intRecordType & " and RecordCount = " & intRecCount & ""
            .Open sSQL
            If .EOF = True And .BOF = True Then
               LogMessage sysMessage, "Record in Scheduel not found RecordType = " & intRecordType & " RecordCount = " & intRecCount
            Else
               Do
                 If .Fields("Extension") = temExten Then
                    .Delete
                    .Update
                 End If
                 .MoveNext
               Loop While .EOF <> True
            End If
            .Close
        End With
        Set ScheduleTable = Nothing
exiterror:
    Exit Sub
OOPS:
    j = j = 1
    LogMessage ErrorMessage, "DeleteScheduleRecord " & Err.Description
     If j <= 2 Then
        If DB.Errors.Count <> 0 Then
           If DB.Errors(0).NativeError <> 0 Then
              If ScheduleTable.State = adStateOpen Then
                 ScheduleTable.Close
                 Set ScheduleTable = Nothing
              End If
              CloseAllTables
              If DB.State = adStateOpen Then
                 DB.Close
                 Set DB = Nothing
              End If
              Delay 500
              
              Resume StartAgain
           End If
        End If
       End If
    Resume exiterror
End Sub
