Imports System.Windows.Forms
Imports System.Reflection

'Namespace ScriptHelp.TaskPane

Public Class Settings

    Public Sub New()
        InitializeComponent()
        Me.pgdSettings.SelectedObject = My.Settings
    End Sub

    Public Shared Sub SetLabelColumnWidth(grid As PropertyGrid, width As Integer)
        If grid Is Nothing Then
            Return
        End If

        Dim fi As FieldInfo = grid.[GetType]().GetField("gridView", BindingFlags.Instance Or BindingFlags.NonPublic)
        If fi Is Nothing Then
            Return
        End If

        Dim view As Control = TryCast(fi.GetValue(grid), Control)
        If view Is Nothing Then
            Return
        End If

        Dim mi As MethodInfo = view.[GetType]().GetMethod("MoveSplitterTo", BindingFlags.Instance Or BindingFlags.NonPublic)
        If mi Is Nothing Then
            Return
        End If
        mi.Invoke(view, New Object() {width})
    End Sub

    Private Sub pgdSettings_PropertyValueChanged(s As Object, e As PropertyValueChangedEventArgs)
        'Scripts.Ribbon.ribbonref.InvalidateRibbon()
    End Sub

End Class

'End Namespace