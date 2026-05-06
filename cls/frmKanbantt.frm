Version =21
VersionRequired =20
PublishOption =1
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8868
    DatasheetFontHeight =11
    ItemSuffix =1
    Right =8868
    Bottom =7488
    TimerInterval =500
    DatasheetGridlinesColor =14806254
    OnUnload ="[Event Procedure]"
    Caption ="Kanban"
    DatasheetFontName ="Calibri"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
        Begin Edge
            OldBorderStyle =1
            BorderColor =10921638
            GridlineColor =10921638
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =7500
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Edge
                    OldBorderStyle =0
                    OverlapFlags =85
                    Width =8868
                    Height =7500
                    Name ="WebBrowser1"
                    HorizontalAnchor =2
                    VerticalAnchor =2

                    LayoutCachedWidth =8868
                    LayoutCachedHeight =7500
                    OnDocumentComplete ="[Event Procedure]"
                    TrustedDomains ="tblTrustedDomains"
                End
            End
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

' -- OOP object graph ----------------------------------------------------------
Private m_Bridge        As clsHtmlBridge
Private m_Router        As clsCommandRouter
Private m_Serializer    As clsDataSerializer
Private m_BoardRepo     As clsBoardRepo
Private m_TaskRepo      As clsTaskRepo
Private m_Processing    As Boolean
Private WithEvents mBroker As clsPubSubBroker
Attribute mBroker.VB_VarHelpID = -1

Private Sub Form_Load()
    ' Wire up the OOP object graph
    Set m_BoardRepo = New clsBoardRepo
    Set m_TaskRepo = New clsTaskRepo
    
    Set m_Bridge = New clsHtmlBridge
    m_Bridge.Initialize Me.WebBrowser1
    
    Dim taskSvc As New clsTaskService
    Set taskSvc.Repo = m_TaskRepo
    
    Dim boardSvc As New clsBoardService
    Set boardSvc.BoardRepo = m_BoardRepo
    
    Set m_Router = New clsCommandRouter
    Set m_Router.TaskService = taskSvc
    Set m_Router.BoardService = boardSvc
    
    Set m_Serializer = New clsDataSerializer
    
    ' -- PubSub broker ---
    Set mBroker = New clsPubSubBroker
    DoCmd.OpenForm "frmPubSub", , , , , acHidden
    Forms!frmPubSub.SetBroker mBroker
    
    PurgeChangeLog 7
    
    DoCmd.Maximize
    Me.WebBrowser1.Navigate "https://msaccess/" & CurrentProject.path & "\view\index.html"
End Sub

Private Sub WebBrowser1_DocumentComplete(URL As Variant)
    If m_Bridge.IsDocumentReady(URL) And Not m_Processing Then
        RefreshUI
    End If
End Sub

Private Sub Form_Timer()
    If m_Processing Then Exit Sub
    On Error GoTo ExitHere
    
    Dim jsonStr As String
    jsonStr = m_Bridge.GetPendingCommand()
    If jsonStr = "" Then GoTo ExitHere
    
    m_Bridge.ClearPendingCommand
    m_Processing = True
    
    Dim restoreBoardId As String
    restoreBoardId = m_Router.Route(jsonStr)
    
    If restoreBoardId <> "NORELOAD" Then
        DoEvents
        RefreshUI restoreBoardId
    End If

ExitHere:
    m_Processing = False
End Sub

Private Sub RefreshUI(Optional ByVal restoreBoardId As String = "")
    If restoreBoardId <> "" Then
        m_Router.ActiveBoardID = restoreBoardId
    End If
    
    Dim boards As Collection: Set boards = m_BoardRepo.GetActive()
    Dim tasks As Collection:  Set tasks = m_TaskRepo.GetAll()
    Dim payload As String:    payload = m_Serializer.BuildPayload(boards, tasks)
    
    m_Bridge.PushData payload, m_Router.ActiveBoardID
End Sub

Private Sub mBroker_TasksChanged(ByVal strPayload As String)
    m_Bridge.ExecuteJS "mergeTaskChanges(" & strPayload & ")"
End Sub

Private Sub mBroker_BoardsChanged(ByVal strPayload As String)
    m_Bridge.ExecuteJS "mergeBoardChanges(" & strPayload & ")"
End Sub

Private Sub mBroker_PollingError(ByVal strMessage As String)
    m_Bridge.ExecuteJS "setSyncStatus(false)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CurrentProject.AllForms("frmPubSub").IsLoaded Then
        DoCmd.Close acForm, "frmPubSub", acSaveNo
    End If
    Set mBroker = Nothing
End Sub
