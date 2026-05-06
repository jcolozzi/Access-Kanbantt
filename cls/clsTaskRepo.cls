Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsTaskRepo — DAO operations for tblTasks
' ============================================================================

' Return a single task by ID, or Nothing if not found.
Public Function GetById(ByVal taskId As Long) As clsTask
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblTasks WHERE TaskID=" & taskId, dbOpenSnapshot)
    If Not rs.EOF Then
        Dim t As clsTask
        Set t = New clsTask
        t.LoadFromRecordset rs
        Set GetById = t
    End If
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function

' Return all non-deleted tasks ordered by SortOrder, TaskID.
Public Function GetAll() As Collection
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim col As New Collection
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblTasks WHERE IsDeleted = False " & _
        "ORDER BY SortOrder, TaskID", dbOpenSnapshot)
    Do Until rs.EOF
        Dim t As clsTask
        Set t = New clsTask
        t.LoadFromRecordset rs
        col.Add t
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set GetAll = col
End Function
' Insert a new task.  Sets CreatedOn/ModifiedOn and auto-assigns SortOrder.
Public Sub Add(task As clsTask)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblTasks", dbOpenDynaset)
    rs.AddNew
    rs!boardId = CStr(task.boardId)
    rs!Title = task.Title
    rs!Description = task.Description
    rs!Priority = IIf(task.Priority = "", "Medium", task.Priority)
    rs!Status = IIf(task.Status = "", "todo", task.Status)
    If Not IsNull(task.StartDate) Then rs!StartDate = task.StartDate
    If Not IsNull(task.EndDate) Then rs!EndDate = task.EndDate
    If Not IsNull(task.StartTime) Then rs!StartTime = task.StartTime
    If Not IsNull(task.EndTime) Then rs!EndTime = task.EndTime
    rs!PctComplete = task.PctComplete
    If task.TaskColor <> "" Then rs!TaskColor = task.TaskColor
    If task.TaskType <> "" Then rs!TaskType = task.TaskType
    If Not IsNull(task.ParentID) Then rs!ParentID = task.ParentID
    rs!Notes = task.Notes
    
    ' Auto-assign SortOrder = max + 1
    Dim maxSort As Variant
    maxSort = DMax("SortOrder", "tblTasks", _
        "BoardID='" & task.boardId & "' AND Status='" & rs!Status & "' AND IsDeleted=False")
    rs!sortOrder = Nz(maxSort, 0) + 1
    
    rs!CreatedOn = Now()
    rs!ModifiedOn = Now()
    rs.Update
    
    rs.Bookmark = rs.LastModified
    Dim lngNewTaskID As Long
    lngNewTaskID = rs!taskId
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    LogChange "task", lngNewTaskID, "add"
End Sub

' Update an existing task by TaskID.
Public Sub Update(task As clsTask)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblTasks WHERE TaskID=" & task.taskId, dbOpenDynaset)
    If rs.EOF Then GoTo CleanUp
    
    rs.Edit
    rs!Status = task.Status
    If task.Title <> "" Then rs!Title = task.Title
    rs!Description = task.Description
    If task.Priority <> "" Then rs!Priority = task.Priority
    If Not IsNull(task.StartDate) Then rs!StartDate = task.StartDate Else rs!StartDate = Null
    If Not IsNull(task.EndDate) Then rs!EndDate = task.EndDate Else rs!EndDate = Null
    If Not IsNull(task.StartTime) Then rs!StartTime = task.StartTime Else rs!StartTime = Null
    If Not IsNull(task.EndTime) Then rs!EndTime = task.EndTime Else rs!EndTime = Null
    rs!PctComplete = task.PctComplete
    If task.TaskColor <> "" Then rs!TaskColor = task.TaskColor Else rs!TaskColor = Null
    If task.TaskType <> "" Then rs!TaskType = task.TaskType Else rs!TaskType = Null
    If Not IsNull(task.ParentID) Then rs!ParentID = task.ParentID Else rs!ParentID = Null
    rs!Notes = task.Notes
    rs!ModifiedOn = Now()
    rs.Update
    
    LogChange "task", task.taskId, "update"

CleanUp:
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

' Soft-delete a task and cascade to its children.
Public Sub SoftDelete(ByVal taskId As Long)
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute "UPDATE tblTasks SET IsDeleted = True WHERE TaskID=" & taskId, dbFailOnError
    db.Execute "UPDATE tblTasks SET IsDeleted = True WHERE ParentID=" & taskId, dbFailOnError
    Set db = Nothing

    LogChange "task", taskId, "delete"
End Sub

' Update a single field by task ID.
Public Sub UpdateField(ByVal taskId As Long, ByVal fieldName As String, ByVal value As Variant)
    Dim db As DAO.Database
    Dim sql As String
    Set db = CurrentDb
    If IsNull(value) Then
        sql = "UPDATE tblTasks SET " & fieldName & " = Null WHERE TaskID=" & taskId
    ElseIf VarType(value) = vbString Then
        sql = "UPDATE tblTasks SET " & fieldName & " = '" & value & "' WHERE TaskID=" & taskId
    ElseIf VarType(value) = vbBoolean Then
        sql = "UPDATE tblTasks SET " & fieldName & " = " & IIf(value, "True", "False") & " WHERE TaskID=" & taskId
    Else
        sql = "UPDATE tblTasks SET " & fieldName & " = " & value & " WHERE TaskID=" & taskId
    End If
    db.Execute sql, dbFailOnError
    Set db = Nothing
    
    LogChange "task", taskId, "update"
End Sub

' Look up the BoardID for a given task.
Public Function GetBoardIdForTask(ByVal taskId As Long) As String
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT BoardID FROM tblTasks WHERE TaskID=" & taskId, dbOpenSnapshot)
    If Not rs.EOF Then GetBoardIdForTask = CStr(rs!boardId)
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function
