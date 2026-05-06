VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataSerializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsDataSerializer — Converts domain objects to JSON payloads
' ============================================================================

' Build the full payload: {"boards":[...],"tasks":[...]}
Public Function BuildPayload(boards As Collection, tasks As Collection) As String
    BuildPayload = "{""boards"":" & BoardsToJson(boards) & _
                   ",""tasks"":" & TasksToJson(tasks) & "}"
End Function

' Serialize a Collection of clsBoard to a JSON array.
Public Function BoardsToJson(boards As Collection) As String
    Dim items As New Collection
    Dim i As Long
    For i = 1 To boards.Count
        items.Add boards(i).ToJson()
    Next i
    BoardsToJson = WrapJsonArray(items)
End Function

' Serialize a Collection of clsTask to a JSON array.
Public Function TasksToJson(tasks As Collection) As String
    Dim items As New Collection
    Dim i As Long
    For i = 1 To tasks.Count
        items.Add tasks(i).ToJson()
    Next i
    TasksToJson = WrapJsonArray(items)
End Function
