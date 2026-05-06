VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsTask — Task Data Transfer Object
' ============================================================================

Public TaskID       As Long
Public BoardID      As Long
Public Title        As String
Public Description  As String
Public Priority     As String
Public Status       As String
Public StartDate    As Variant      ' Date or Null
Public EndDate      As Variant      ' Date or Null
Public StartTime    As Variant      ' Date or Null
Public EndTime      As Variant      ' Date or Null
Public PctComplete  As Integer
Public TaskColor    As String
Public TaskType     As String
Public ParentID     As Variant      ' Long or Null
Public Notes        As String
Public IsExpanded   As Boolean
Public SortOrder    As Long
Public IsDeleted    As Boolean

Private Sub Class_Initialize()
    Priority = "Medium"
    Status = "todo"
    IsExpanded = True
End Sub

' Populate all properties from a DAO Recordset row (cursor must be on the row).
Public Sub LoadFromRecordset(rs As DAO.Recordset)
    TaskID = Nz(rs!TaskID, 0)
    BoardID = Nz(rs!BoardID, 0)
    Title = SafeString(rs!Title)
    Description = SafeString(rs!Description)
    Priority = IIf(Nz(rs!Priority, "") = "", "Medium", rs!Priority)
    Status = IIf(Nz(rs!Status, "") = "", "todo", rs!Status)
    StartDate = rs!StartDate
    EndDate = rs!EndDate
    StartTime = rs!StartTime
    EndTime = rs!EndTime
    PctComplete = Nz(rs!PctComplete, 0)
    TaskColor = SafeString(rs!TaskColor)
    TaskType = SafeString(rs!TaskType)
    ParentID = rs!ParentID
    Notes = SafeString(rs!Notes)
    IsExpanded = Nz(rs!IsExpanded, True)
    SortOrder = Nz(rs!SortOrder, 0)
    IsDeleted = Nz(rs!IsDeleted, False)
End Sub

' Populate all properties from a JSON payload string.
Public Sub LoadFromJson(ByVal jsonStr As String)
    Dim v As String
    
    v = ExtractJSONValue(jsonStr, "id")
    If v <> "" Then TaskID = SafeLong(v)
    
    v = ExtractJSONValue(jsonStr, "boardId")
    If v <> "" Then BoardID = SafeLong(v)
    
    Title = ExtractJSONValue(jsonStr, "title")
    Description = ExtractJSONValue(jsonStr, "desc")
    
    v = ExtractJSONValue(jsonStr, "priority")
    Priority = IIf(v = "", "Medium", v)
    
    v = ExtractJSONValue(jsonStr, "status")
    Status = IIf(v = "", "todo", v)
    
    v = ExtractJSONValue(jsonStr, "startDate")
    If v <> "" Then StartDate = CDate(v) Else StartDate = Null
    
    v = ExtractJSONValue(jsonStr, "endDate")
    If v <> "" Then EndDate = CDate(v) Else EndDate = Null
    
    v = ExtractJSONValue(jsonStr, "startTime")
    If v <> "" Then StartTime = CDate(v) Else StartTime = Null
    
    v = ExtractJSONValue(jsonStr, "endTime")
    If v <> "" Then EndTime = CDate(v) Else EndTime = Null
    
    v = ExtractJSONValue(jsonStr, "pctComplete")
    PctComplete = SafeInt(v)
    
    TaskColor = ExtractJSONValue(jsonStr, "taskColor")
    TaskType = ExtractJSONValue(jsonStr, "taskType")
    
    v = ExtractJSONValue(jsonStr, "parentId")
    If v <> "" And v <> "0" Then ParentID = SafeLong(v) Else ParentID = Null
    
    Notes = ExtractJSONValue(jsonStr, "notes")
    
    v = ExtractJSONValue(jsonStr, "isExpanded")
    IsExpanded = (v <> "false")
    
    v = ExtractJSONValue(jsonStr, "sortOrder")
    SortOrder = SafeLong(v)
End Sub

' Serialize this task to a JSON object string.
Public Function ToJson() As String
    Dim c As New Collection
    c.Add BuildJsonPair("id", CStr(TaskID))
    c.Add BuildJsonPair("boardId", CStr(BoardID))
    c.Add BuildJsonPair("title", Title)
    c.Add BuildJsonPair("desc", Description)
    c.Add BuildJsonPair("priority", Priority)
    c.Add BuildJsonPair("status", Status)
    c.Add BuildJsonPair("startDate", FormatDateISO(StartDate))
    c.Add BuildJsonPair("endDate", FormatDateISO(EndDate))
    c.Add BuildJsonPair("startTime", FormatTimeHM(StartTime))
    c.Add BuildJsonPair("endTime", FormatTimeHM(EndTime))
    c.Add BuildJsonPair("pctComplete", CStr(PctComplete), False)
    c.Add BuildJsonPair("taskColor", TaskColor)
    c.Add BuildJsonPair("taskType", TaskType)
    c.Add BuildJsonPair("parentId", SafeString(ParentID))
    c.Add BuildJsonPair("notes", Notes)
    c.Add BuildJsonPair("isExpanded", IIf(IsExpanded, "true", "false"), False)
    c.Add BuildJsonPair("sortOrder", CStr(SortOrder), False)
    
    Dim result As String
    Dim i As Long
    result = "{"
    For i = 1 To c.Count
        If i > 1 Then result = result & ","
        result = result & c(i)
    Next i
    ToJson = result & "}"
End Function
