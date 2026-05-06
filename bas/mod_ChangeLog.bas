Option Compare Database
Option Explicit

' Logs a change to tblChangeLog
Public Sub LogChange(ByVal strChangeType As String, ByVal lngRecordId As Long, ByVal strAction As String)
    On Error Resume Next
    Dim strSQL As String
    strSQL = "INSERT INTO tblChangeLog ([ChangeType], [RecordID], [Action], [ChangedBy], [ChangedOn]) " & _
             "VALUES ('" & Replace(strChangeType, "'", "''") & "', " & lngRecordId & ", '" & Replace(strAction, "'", "''") & "', '" & Replace(Environ$("USERNAME"), "'", "''") & "', '" & Format$(Now(), "yyyy-mm-dd hh:nn:ss") & "')"
    CurrentDb.Execute strSQL, dbFailOnError
End Sub

' Purges old change log entries
Public Sub PurgeChangeLog(Optional ByVal lngRetainDays As Long = 7)
    On Error Resume Next
    Dim strSQL As String
    strSQL = "DELETE FROM tblChangeLog WHERE [ChangedOn] < " & _
             "#" & Format$(DateAdd("d", -lngRetainDays, Now()), "yyyy-mm-dd hh:nn:ss") & "#"
    CurrentDb.Execute strSQL, dbFailOnError
End Sub
