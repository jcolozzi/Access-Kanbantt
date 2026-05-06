VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommandRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsCommandRouter — Dispatches JSON commands to service methods
' ============================================================================

Public TaskService  As clsTaskService
Public BoardService As clsBoardService
Public ActiveBoardID As String

' Route a JSON command to the appropriate service.
' Returns a boardId to restore, or "NORELOAD" to skip data refresh.
Public Function Route(ByVal jsonStr As String) As String
    Dim action As String
    action = ExtractJSONValue(jsonStr, "action")
    
    Select Case action
    
        Case "setActiveBoard"
            Dim sbId As String
            sbId = ExtractJSONValue(jsonStr, "boardId")
            If sbId <> "" Then ActiveBoardID = sbId
            Route = "NORELOAD"
        
        Case "addBoard"
            BoardService.AddBoard jsonStr
            Route = ActiveBoardID
        
        Case "renameBoard"
            BoardService.RenameBoard jsonStr
            Route = "NORELOAD"
        
        Case "updateBoard"
            BoardService.UpdateColor jsonStr
            Route = "NORELOAD"
        
        Case "deleteBoard"
            Dim firstId As String
            firstId = BoardService.DeleteBoard(jsonStr)
            ActiveBoardID = firstId
            Route = firstId
        
        Case "addTask"
            Route = TaskService.AddTask(jsonStr)
        
        Case "updateTask"
            Route = TaskService.UpdateTask(jsonStr)
        
        Case "deleteTask"
            Route = TaskService.DeleteTask(jsonStr)
        
        Case "toggleExpand"
            Route = TaskService.ToggleExpand(jsonStr)
        
        Case "reorderTask"
            TaskService.ReorderTask jsonStr
            Route = "NORELOAD"
        
        Case "updateAccent"
            Route = "NORELOAD"
        
        Case Else
            Route = "NORELOAD"
    
    End Select
End Function
