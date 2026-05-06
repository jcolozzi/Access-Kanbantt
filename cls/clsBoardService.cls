VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBoardService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsBoardService — Board business logic
' ============================================================================

Public BoardRepo As clsBoardRepo

' Add a board from a JSON payload.
Public Sub AddBoard(ByVal jsonStr As String)
    Dim b As New clsBoard
    b.LoadFromJson jsonStr
    BoardRepo.Add b
End Sub

' Rename a board from a JSON payload.
Public Sub RenameBoard(ByVal jsonStr As String)
    Dim boardId As Long
    Dim newName As String
    boardId = SafeLong(ExtractJSONValue(jsonStr, "id"))
    newName = ExtractJSONValue(jsonStr, "name")
    If boardId > 0 And newName <> "" Then
        BoardRepo.Rename boardId, newName
    End If
End Sub

' Update a board's color from a JSON payload.
Public Sub UpdateColor(ByVal jsonStr As String)
    Dim boardId As Long
    Dim color As String
    boardId = SafeLong(ExtractJSONValue(jsonStr, "id"))
    color = ExtractJSONValue(jsonStr, "color")
    If boardId > 0 Then
        BoardRepo.UpdateColor boardId, color
    End If
End Sub

' Soft-delete a board and its tasks. Returns the first remaining board ID.
Public Function DeleteBoard(ByVal jsonStr As String) As String
    Dim boardId As Long
    boardId = SafeLong(ExtractJSONValue(jsonStr, "id"))
    BoardRepo.SoftDelete boardId
    DeleteBoard = BoardRepo.GetFirstActiveBoardId()
End Function
