Attribute VB_Name = "mod_JsonUtility"
Option Compare Database
Option Explicit

' ============================================================================
' mod_JsonUtility — Stateless JSON helper functions
' ============================================================================

' Escape a string for safe embedding inside a JSON string value.
Public Function JSONEscape(ByVal v As Variant) As String
    If IsNull(v) Then JSONEscape = "": Exit Function
    Dim s As String
    s = CStr(v)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JSONEscape = s
End Function

' Extract a value from flat JSON by key name.
' Handles both "key":"stringval" and "key":numericval
Public Function ExtractJSONValue(ByVal jsonStr As String, ByVal key As String) As String
    Dim searchKey As String
    Dim p As Long, startPos As Long, endPos As Long
    searchKey = """" & key & """:"
    p = InStr(1, jsonStr, searchKey, vbTextCompare)
    If p = 0 Then Exit Function
    startPos = p + Len(searchKey)
    If Mid$(jsonStr, startPos, 1) = """" Then
        startPos = startPos + 1
        endPos = InStr(startPos, jsonStr, """")
        ExtractJSONValue = Mid$(jsonStr, startPos, endPos - startPos)
    Else
        endPos = InStr(startPos, jsonStr, ",")
        If endPos = 0 Then endPos = InStr(startPos, jsonStr, "}")
        ExtractJSONValue = Trim$(Mid$(jsonStr, startPos, endPos - startPos))
    End If
End Function

' Build a single JSON key:value pair.
' If asString=True:  "key":"escapedValue"
' If asString=False: "key":rawValue
Public Function BuildJsonPair(ByVal key As String, ByVal value As String, _
                               Optional ByVal asString As Boolean = True) As String
    If asString Then
        BuildJsonPair = """" & key & """:""" & JSONEscape(value) & """"
    Else
        BuildJsonPair = """" & key & """:" & value
    End If
End Function

' Wrap comma-separated pairs into a JSON object string: {pair1,pair2,...}
Public Function WrapJsonObject(ParamArray pairs() As Variant) As String
    Dim i As Long
    Dim result As String
    result = "{"
    For i = LBound(pairs) To UBound(pairs)
        If i > LBound(pairs) Then result = result & ","
        result = result & CStr(pairs(i))
    Next i
    WrapJsonObject = result & "}"
End Function

' Join a Collection of JSON object strings into a JSON array: [obj1,obj2,...]
Public Function WrapJsonArray(items As Collection) As String
    Dim result As String
    Dim i As Long
    result = "["
    For i = 1 To items.Count
        If i > 1 Then result = result & ","
        result = result & items(i)
    Next i
    WrapJsonArray = result & "]"
End Function
