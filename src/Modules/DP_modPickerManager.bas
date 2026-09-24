Attribute VB_Name = "DP_modPickerManager"
Option Explicit

'----------------------------------------
' Declarations
'----------------------------------------

Private PickerForms As Collection


'----------------------------------------
' Retrieval
'----------------------------------------

Public Function DP_GetPickerForm(ByVal TargetWindow As Excel.Window) As DP_frmDatePicker

    Dim Picker As DP_frmDatePicker

    If PickerForms Is Nothing Then
        Set PickerForms = New Collection
    End If

    Set Picker = FindPickerForm(TargetWindow)

    If Picker Is Nothing Then
        Set Picker = New DP_frmDatePicker
        Set Picker.TargetWindow = TargetWindow

        PickerForms.Add Picker
    End If

    Set DP_GetPickerForm = Picker

End Function


Private Function FindPickerForm(ByVal TargetWindow As Excel.Window) As DP_frmDatePicker

    Dim Index As Long
    Dim Picker As DP_frmDatePicker
    Dim TargetWindowKey As String

    If PickerForms Is Nothing Then Exit Function

    TargetWindowKey = DP_GetWindowKey(TargetWindow)

    For Index = PickerForms.Count To 1 Step -1

        Set Picker = GetPickerFormAt(Index)

        If Not Picker Is Nothing Then
            If DP_GetWindowKey(Picker.TargetWindow) = TargetWindowKey Then
                Set FindPickerForm = Picker
                Exit Function
            End If
        End If

    Next Index

End Function


Private Function GetPickerFormAt(ByVal Index As Long) As DP_frmDatePicker

    Dim Picker As DP_frmDatePicker

    ' UserForm_Terminate normally removes picker instances.
    ' This is a defensive fallback for stale references.
    On Error Resume Next
    Set Picker = PickerForms(Index)
    On Error GoTo 0

    If Picker Is Nothing Then
        PickerForms.Remove Index
    Else
        Set GetPickerFormAt = Picker
    End If

End Function


'----------------------------------------
' Lifecycle
'----------------------------------------

Public Sub DP_RemovePickerForm(ByVal Picker As DP_frmDatePicker)

    Dim Index As Long
    Dim CurrentPicker As DP_frmDatePicker

    If PickerForms Is Nothing Then Exit Sub

    For Index = PickerForms.Count To 1 Step -1

        Set CurrentPicker = GetPickerFormAt(Index)

        If Not CurrentPicker Is Nothing Then
            If CurrentPicker Is Picker Then
                PickerForms.Remove Index
                Exit Sub
            End If
        End If

    Next Index

End Sub


Public Sub DP_HidePickerForm()

    Dim Picker As DP_frmDatePicker

    If PickerForms Is Nothing Then Exit Sub

    Set Picker = FindPickerForm(Excel.Application.ActiveWindow)

    If Picker Is Nothing Then Exit Sub

    If Picker.Visible Then
        Picker.Hide
    End If

End Sub

