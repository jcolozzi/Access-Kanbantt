Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Event: TasksChanged
Public Event TasksChanged(ByVal strJsonPayload As String)
' Event: BoardsChanged
Public Event BoardsChanged(ByVal strJsonPayload As String)
' Event: PollingError
Public Event PollingError(ByVal strErrDesc As String)

' Publishes a TasksChanged event
Public Sub PublishTasksChanged(ByVal strJsonPayload As String)
    RaiseEvent TasksChanged(strJsonPayload)
End Sub

' Publishes a BoardsChanged event
Public Sub PublishBoardsChanged(ByVal strJsonPayload As String)
    RaiseEvent BoardsChanged(strJsonPayload)
End Sub

' Publishes a PollingError event
Public Sub PublishError(ByVal strErrDesc As String)
    RaiseEvent PollingError(strErrDesc)
End Sub
