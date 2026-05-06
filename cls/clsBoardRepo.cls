Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsBoardRepo — DAO operations for tblBoards
' ============================================================================

' Return all active (non-deleted) boards.
Public Function GetActive() As Collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim col As New Collection
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblBoards WHERE IsDeleted = False ORDER BY BoardID", dbOpenSnapshot)
    Do Until rs.EOF
        Dim b As clsBoard
        Set b = New clsBoard
        b.LoadFromRecordset rs
        col.Add b
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set GetActive = col
End Function

' Add a new board.
Public Sub Add(board As clsBoard)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblBoards", dbOpenDynaset)
    rs.AddNew
    rs!BoardName = board.BoardName
    If board.BoardColor <> "" Then rs!BoardColor = board.BoardColor
    rs.Update
    
    Dim lngNewBoardID As Long
    lngNewBoardID = rs!boardId
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    LogChange "board", lngNewBoardID, "add"
End Sub

' Rename a board.
Public Sub Rename(ByVal boardId As Long, ByVal newName As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblBoards WHERE BoardID=" & boardId, dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        rs!BoardName = newName
        rs.Update
        LogChange "board", boardId, "update"
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' Update a board's color.
Public Sub UpdateColor(ByVal boardId As Long, ByVal color As String)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblBoards WHERE BoardID=" & boardId, dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        If color <> "" Then rs!BoardColor = color Else rs!BoardColor = Null
        rs.Update
        LogChange "board", boardId, "update"
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' Soft-delete a board and all its tasks.
Public Sub SoftDelete(ByVal boardId As Long)
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "UPDATE tblBoards SET IsDeleted = True WHERE BoardID = " & boardId, dbFailOnError
    db.Execute "UPDATE tblTasks  SET IsDeleted = True WHERE BoardID = " & boardId, dbFailOnError
    Set db = Nothing

    LogChange "board", boardId, "delete"
End Sub

' Return the ID of the first active board, or "" if none exist.
Public Function GetFirstActiveBoardId() As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT MIN(BoardID) AS MinID FROM tblBoards WHERE IsDeleted = False", dbOpenSnapshot)
    If Not rs.EOF Then GetFirstActiveBoardId = SafeString(rs!MinID)
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function
