VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsBoard — Board Data Transfer Object
' ============================================================================

Public BoardID      As Long
Public BoardName    As String
Public BoardColor   As String
Public CreatedOn    As Variant      ' Date or Null
Public IsDeleted    As Boolean

' Populate from a DAO Recordset row.
Public Sub LoadFromRecordset(rs As DAO.Recordset)
    BoardID = Nz(rs!BoardID, 0)
    BoardName = SafeString(rs!BoardName)
    BoardColor = SafeString(rs!BoardColor)
    On Error Resume Next
    CreatedOn = rs!CreatedOn
    On Error GoTo 0
    IsDeleted = Nz(rs!IsDeleted, False)
End Sub

' Populate from a JSON payload string.
Public Sub LoadFromJson(ByVal jsonStr As String)
    Dim v As String
    v = ExtractJSONValue(jsonStr, "id")
    If v <> "" Then BoardID = SafeLong(v)
    BoardName = ExtractJSONValue(jsonStr, "name")
    BoardColor = ExtractJSONValue(jsonStr, "color")
End Sub

' Serialize to a JSON object string.
Public Function ToJson() As String
    ToJson = WrapJsonObject( _
        BuildJsonPair("id", CStr(BoardID)), _
        BuildJsonPair("name", BoardName), _
        BuildJsonPair("color", BoardColor))
End Function
