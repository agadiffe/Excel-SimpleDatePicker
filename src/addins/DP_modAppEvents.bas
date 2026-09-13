Attribute VB_Name = "DP_modAppEvents"
Option Explicit

Private AppEvents As DP_CAppEvents


Public Sub DP_InitializeAppEvents()

    Set AppEvents = New DP_CAppEvents
    Set AppEvents.App = Excel.Application

End Sub

