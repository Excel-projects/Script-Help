Imports System.Windows.Forms
Imports System.Data.SqlServerCe
Imports ScriptHelp.Scripts

Public Class TableData

    Public Sub New()
        InitializeComponent()
        Try
            dgvList.AutoGenerateColumns = True
            Dim tableName As String = Ribbon.AppVariables.TableName
            Me.Text = "List of " + tableName
            Select Case tableName
                Case "TableAlias"
                    dgvList.DataSource = Data.TableAliasTable
                    Exit Select
                Case "DateFormat"
                    dgvList.DataSource = Data.DateFormatTable
                    Exit Select
            End Select

            Me.dgvList.Columns(0).Width = dgvList.Width - 75

        Catch ex As Exception
            'ErrorHandler.DisplayMessage(ex)

        End Try

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Me.Validate()
            If dgvList.IsCurrentRowDirty OrElse dgvList.IsCurrentCellDirty Then
                dgvList.CommitEdit(DataGridViewDataErrorContexts.Commit)
                dgvList.EndEdit()
            End If

            Dim tableName As String = Ribbon.AppVariables.TableName
            'String sql = "SELECT * FROM @tableName";
            Dim sql As String = Convert.ToString("SELECT * FROM ") & tableName
            Dim cn As New SqlCeConnection(Data.Connection())
            Dim scb As SqlCeCommandBuilder = Nothing
            Dim sda As New SqlCeDataAdapter(sql, cn)

            sda.TableMappings.Add("Table", tableName)
            scb = New SqlCeCommandBuilder(sda)
            Select Case tableName
                Case "TableAlias"
                    sda.Update(Data.TableAliasTable)
                    Data.CreateTableAliasTable()
                    Exit Select
                Case "DateFormat"
                    sda.Update(Data.DateFormatTable)
                    Data.CreateDateFormatTable()
                    Exit Select
            End Select

            'Ribbon.ribbonref.InvalidateRibbon()

        Catch ex As Exception
            'ErrorHandler.DisplayMessage(ex)

        End Try

    End Sub

End Class
