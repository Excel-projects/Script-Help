Imports System.Windows.Forms

'Namespace ScriptHelp.TaskPane

Public Class Script

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
        Try
            Me.txtScript.SelectAll()
            Me.txtScript.Copy()

        Catch ex As Exception
            'ErrorHandler.DisplayMessage(ex)
        End Try

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim s As New SaveFileDialog()
            Select Case Ribbon.AppVariables.FileType
                Case "SQL"
                    s.FileName = "Update_" + My.Settings.Table_ColumnTableAlias + ".sql"
                    s.Filter = "Structured Query Language | *.sql"
                    Exit Select
                Case "DQL"
                    s.FileName = "Update_" + Ribbon.AppVariables.FirstColumnName + ".dql"
                    s.Filter = "Documentum Query Language | *.dql"
                    Exit Select
                Case "TXT"
                    s.FileName = My.Settings.Table_ColumnTableAlias + ".txt"
                    s.Filter = "Text File | *.txt"
                    Exit Select
                Case "XML"
                    s.FileName = My.Settings.Table_ColumnTableAlias + ".xml"
                    s.Filter = "Extensible Markup Language | *.xml"
                    Exit Select
            End Select
            If s.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Using sw As New System.IO.StreamWriter(s.FileName)
                    For Each line As String In txtScript.Lines
                        sw.WriteLine(line)
                    Next
                End Using
            End If

        Catch ex As Exception
            'ErrorHandler.DisplayMessage(ex)
        End Try

    End Sub

    Private Sub Script_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtScript.Text = Ribbon.AppVariables.ScriptRange
    End Sub

End Class

'End Namespace