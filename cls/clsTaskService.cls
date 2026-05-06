VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsTaskService — Task business logic
' ============================================================================

Public Repo As clsTaskRepo

' Add a task from a JSON payload. Returns the boardId.
Public Function AddTask(ByVal jsonStr As String) As String
    Dim t As New clsTask
    t.LoadFromJson jsonStr
    Repo.Add t
    AddTask = CStr(t.BoardID)
End Function

' Update a task from a JSON payload. Returns the boardId.
Public Function UpdateTask(ByVal jsonStr As String) As String
    Dim t As New clsTask
    t.LoadFromJson jsonStr
    ' Look up the boardId from DB (authoritative source)
    UpdateTask = Repo.GetBoardIdForTask(t.TaskID)
    Repo.Update t
End Function

' Soft-delete a task. Returns the boardId.
Public Function DeleteTask(ByVal jsonStr As String) As String
    Dim taskId As Long
    taskId = SafeLong(ExtractJSONValue(jsonStr, "id"))
    DeleteTask = Repo.GetBoardIdForTask(taskId)
    Repo.SoftDelete taskId
End Function

' Toggle the IsExpanded flag on a task. Returns the boardId.
Public Function ToggleExpand(ByVal jsonStr As String) As String
    Dim taskId As Long
    Dim expanded As String
    taskId = SafeLong(ExtractJSONValue(jsonStr, "id"))
    expanded = ExtractJSONValue(jsonStr, "isExpanded")
    ToggleExpand = Repo.GetBoardIdForTask(taskId)
    Repo.UpdateField taskId, "IsExpanded", (expanded = "true")
End Function

' Update a task's sort order. Returns "NORELOAD".
Public Sub ReorderTask(ByVal jsonStr As String)
    Dim taskId As Long
    Dim sortOrder As Long
    taskId = SafeLong(ExtractJSONValue(jsonStr, "id"))
    sortOrder = SafeLong(ExtractJSONValue(jsonStr, "sortOrder"))
    Repo.UpdateField taskId, "SortOrder", sortOrder
End Sub
