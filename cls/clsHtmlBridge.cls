VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHtmlBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ============================================================================
' clsHtmlBridge — Abstraction over WebBrowser control for JS <-> VBA comms
' ============================================================================

Private m_WebBrowser As Object

' Store a reference to the form's WebBrowser control.
Public Sub Initialize(wb As Object)
    Set m_WebBrowser = wb
End Sub

' Execute arbitrary JavaScript on the page.
Public Sub ExecuteJS(ByVal code As String)
    If m_WebBrowser Is Nothing Then Exit Sub
    m_WebBrowser.ExecuteJavascript code
End Sub

' Retrieve a JavaScript value as a string.
Public Function RetrieveValue(ByVal jsExpr As String) As String
    If m_WebBrowser Is Nothing Then RetrieveValue = "": Exit Function
    RetrieveValue = Nz(m_WebBrowser.RetrieveJavascriptValue(jsExpr), "")
End Function

' Get the pending command from JavaScript, with JSON unescaping.
' Returns "" if nothing pending.
Public Function GetPendingCommand() As String
    Dim raw As String
    raw = RetrieveValue("getPendingCommand()")
    If raw = "" Or raw = "null" Then
        GetPendingCommand = ""
        Exit Function
    End If
    
    ' Unwrap outer quotes and unescape
    If Left$(raw, 1) = """" And Right$(raw, 1) = """" Then
        raw = Mid$(raw, 2, Len(raw) - 2)
        raw = Replace(raw, "\""", """")
        raw = Replace(raw, "\\", "\")
    End If
    
    GetPendingCommand = raw
End Function

' Clear the pending command in JavaScript.
Public Sub ClearPendingCommand()
    ExecuteJS "clearPendingCommand();"
End Sub

' Push the full data payload to the page and optionally switch to a board.
Public Sub PushData(ByVal jsonPayload As String, Optional ByVal restoreBoardId As String = "")
    Dim restoreCmd As String
    If restoreBoardId <> "" Then
        restoreCmd = " window.switchBoard('" & restoreBoardId & "');"
    End If
    ExecuteJS "window.loadData(JSON.stringify(" & jsonPayload & "));" & restoreCmd
End Sub

' Check if the document-complete URL is for our main page.
Public Function IsDocumentReady(ByVal url As Variant) As Boolean
    IsDocumentReady = (InStr(1, CStr(url), "index.html", vbTextCompare) > 0)
End Function
