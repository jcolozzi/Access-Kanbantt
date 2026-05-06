Attribute VB_Name = "mod_TypeConverter"
Option Compare Database
Option Explicit

' ============================================================================
' mod_TypeConverter — Safe type conversions and date/time formatting
' ============================================================================

' Convert any Variant to Long; returns 0 on error or Null/Empty.
Public Function SafeLong(ByVal v As Variant) As Long
    On Error Resume Next
    SafeLong = 0
    If Not IsNull(v) And Not IsEmpty(v) Then
        If Len(Trim$(CStr(v))) > 0 Then SafeLong = CLng(v)
    End If
    On Error GoTo 0
End Function

' Convert any Variant to Integer; returns 0 on error or Null/Empty.
Public Function SafeInt(ByVal v As Variant) As Integer
    On Error Resume Next
    SafeInt = 0
    If Not IsNull(v) And Not IsEmpty(v) Then
        If Len(Trim$(CStr(v))) > 0 Then SafeInt = CInt(v)
    End If
    On Error GoTo 0
End Function

' Convert any Variant to Date; returns today on error or Null/Empty.
Public Function SafeDate(ByVal v As Variant) As Date
    On Error Resume Next
    SafeDate = Date
    If Not IsNull(v) And Not IsEmpty(v) Then
        If IsDate(v) Then SafeDate = CDate(v)
    End If
    On Error GoTo 0
End Function

' Convert any Variant to String; returns "" on Null/Empty.
Public Function SafeString(ByVal v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        SafeString = ""
    Else
        SafeString = CStr(v)
    End If
End Function

' Format a Date/Variant to ISO date string "yyyy-mm-dd"; returns "" if Null.
Public Function FormatDateISO(ByVal d As Variant) As String
    If IsNull(d) Then
        FormatDateISO = ""
    Else
        FormatDateISO = Format$(d, "yyyy-mm-dd")
    End If
End Function

' Format a Date/Variant to time string "hh:nn"; returns "" if Null.
Public Function FormatTimeHM(ByVal t As Variant) As String
    If IsNull(t) Then
        FormatTimeHM = ""
    Else
        FormatTimeHM = Format$(t, "hh:nn")
    End If
End Function
