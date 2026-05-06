Version =21
VersionRequired =20
PublishOption =1
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    TimerInterval =3000
    DatasheetGridlinesColor =14806254
    Caption ="PubSub Poller"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Section
            Height =7560
            Name ="Detail"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private mlngLastChangeID As Long
Private mstrCurrentUser As String
Private mblnProcessing As Boolean
Private mBroker As clsPubSubBroker
Private mTaskRepo As clsTaskRepo
Private mBoardRepo As clsBoardRepo

Public Sub SetBroker(ByVal oBroker As clsPubSubBroker)
    Set mBroker = oBroker
    Set mTaskRepo = New clsTaskRepo
    Set mBoardRepo = New clsBoardRepo
End Sub

Private Sub Form_Load()
    mstrCurrentUser = Environ$("USERNAME")
    mlngLastChangeID = Nz(DMax("ChangeID", "tblChangeLog"), 0)
End Sub

Private Sub Form_Timer()
    If mblnProcessing Then Exit Sub
    If mBroker Is Nothing Then Exit Sub
    mblnProcessing = True
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    strSQL = "SELECT ChangeID, ChangeType, RecordID, [Action] FROM tblChangeLog " & _
        "WHERE ChangeID > " & mlngLastChangeID & _
        " AND ChangedBy <> '" & Replace$(mstrCurrentUser, "'", "''") & "'" & _
        " ORDER BY ChangeID"
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    If rs.EOF Then
        rs.Close: Set rs = Nothing: Set db = Nothing
        mblnProcessing = False
        Exit Sub
    End If
    
    Dim lngMaxID As Long
    lngMaxID = mlngLastChangeID
    
    Dim colTaskItems As New Collection
    Dim colBoardItems As New Collection
    Dim strType As String
    Dim lngRecID As Long
    Dim strAct As String
    Dim strItem As String
    
    Do Until rs.EOF
        strType = Nz(rs!ChangeType, "")
        lngRecID = Nz(rs!RecordID, 0)
        strAct = Nz(rs![action], "")
        
        If rs!ChangeID > lngMaxID Then lngMaxID = rs!ChangeID
        
        If strAct = "delete" Then
            strItem = "{" & Chr$(34) & "action" & Chr$(34) & ":" & Chr$(34) & "delete" & Chr$(34) & "," & Chr$(34) & "id" & Chr$(34) & ":" & lngRecID & "}"
        Else
            strItem = ""
            If strType = "task" Then
                Dim oTask As clsTask
                Set oTask = mTaskRepo.GetById(lngRecID)
                If Not oTask Is Nothing Then
                    strItem = "{" & Chr$(34) & "action" & Chr$(34) & ":" & Chr$(34) & strAct & Chr$(34) & "," & Chr$(34) & "data" & Chr$(34) & ":" & oTask.ToJson() & "}"
                End If
            ElseIf strType = "board" Then
                Dim oBrd As clsBoard
                Set oBrd = GetBoardById(lngRecID)
                If Not oBrd Is Nothing Then
                    strItem = "{" & Chr$(34) & "action" & Chr$(34) & ":" & Chr$(34) & strAct & Chr$(34) & "," & Chr$(34) & "data" & Chr$(34) & ":" & oBrd.ToJson() & "}"
                End If
            End If
        End If
        
        If strItem <> "" Then
            If strType = "task" Then
                colTaskItems.Add strItem
            ElseIf strType = "board" Then
                colBoardItems.Add strItem
            End If
        End If
        
        rs.MoveNext
    Loop
    rs.Close: Set rs = Nothing: Set db = Nothing
    
    mlngLastChangeID = lngMaxID
    
    If colTaskItems.Count > 0 Then
        mBroker.PublishTasksChanged WrapJsonArray(colTaskItems)
    End If
    If colBoardItems.Count > 0 Then
        mBroker.PublishBoardsChanged WrapJsonArray(colBoardItems)
    End If
    
    mblnProcessing = False
    Exit Sub
    
ErrHandler:
    Me.TimerInterval = 0
    If Not mBroker Is Nothing Then
        mBroker.PublishError Err.Description
    End If
    mblnProcessing = False
End Sub

Private Function GetBoardById(ByVal lngBoardId As Long) As clsBoard
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset( _
        "SELECT * FROM tblBoards WHERE BoardID=" & lngBoardId, dbOpenSnapshot)
    If Not rs.EOF Then
        Dim b As clsBoard
        Set b = New clsBoard
        b.LoadFromRecordset rs
        Set GetBoardById = b
    End If
    rs.Close: Set rs = Nothing: Set db = Nothing
End Function

Private Function WrapJsonArray(ByVal col As Collection) As String
    Dim strOut As String
    Dim i As Long
    strOut = "["
    For i = 1 To col.Count
        If i > 1 Then strOut = strOut & ","
        strOut = strOut & col(i)
    Next i
    strOut = strOut & "]"
    WrapJsonArray = strOut
End Function
