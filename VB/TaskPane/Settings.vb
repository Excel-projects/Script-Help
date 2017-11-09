Imports System.Windows.Forms
Imports System.Reflection

'Namespace ScriptHelp.TaskPane
Imports ScriptHelp.Scripts
Imports ScriptHelp.Ribbon

Public Class Settings

    Public Sub New()
        Try
            InitializeComponent()
            Me.pgdSettings.SelectedObject = My.Settings

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)

        End Try

    End Sub

    Public Shared Sub SetLabelColumnWidth(grid As PropertyGrid, width As Integer)
        Try
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

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)

        End Try

    End Sub

    Private Sub pgdSettings_PropertyValueChanged(s As Object, e As PropertyValueChangedEventArgs) Handles pgdSettings.PropertyValueChanged
        Try
            ribbonref.InvalidateRibbon()

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)

        End Try

    End Sub

End Class

'End Namespace