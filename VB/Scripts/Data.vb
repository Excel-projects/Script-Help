Imports System
Imports System.IO
Imports System.Data
Imports System.Data.SqlServerCe
Imports System.Security.AccessControl
Imports System.Windows.Forms

Namespace Scripts

    Public Class Data

        Const dataFolder As String = "App_Data"

        Public Shared Function Connection() As String
            Data.SetUserPath()
            Dim databaseFile As String = "Data Source=" + Path.Combine(My.Settings.App_PathLocalData, My.Application.Info.ProductName + ".sdf")
            Return databaseFile
        End Function

        Public Shared TableAliasTable As New DataTable()
        Public Shared DateFormatTable As New DataTable()
        Public Shared TimeFormatTable As New DataTable()
        Public Shared GraphDataTable As New DataTable()

        Public Shared Sub CreateTableAliasTable()
            Try
                Dim columnName As String = "TableName"
                Dim dcTableName As New DataColumn(columnName, GetType(String))
                TableAliasTable.Rows.Clear()
                Dim columns As DataColumnCollection = TableAliasTable.Columns
                If columns.Contains(columnName) = False Then
                    TableAliasTable.Columns.Add(dcTableName)
                End If

                Dim tableName As String = "TableAlias"
                Dim sql As String = Convert.ToString((Convert.ToString("SELECT * FROM ") & tableName) + " ORDER BY ") & columnName

                Using da = New SqlCeDataAdapter(sql, Connection())
                    da.Fill(TableAliasTable)
                End Using

                TableAliasTable.DefaultView.Sort = columnName & Convert.ToString(" asc")
                TableAliasTable.TableName = tableName

            Catch ex As Exception
                'ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub CreateDateFormatTable()
            Try
                Dim columnName As String = "FormatString"
                Dim dcFormatString As New DataColumn(columnName, GetType(String))
                DateFormatTable.Rows.Clear()
                Dim columns As DataColumnCollection = DateFormatTable.Columns
                If columns.Contains(columnName) = False Then
                    DateFormatTable.Columns.Add(dcFormatString)
                End If

                Dim tableName As String = "DateFormat"
                Dim sql As String = Convert.ToString((Convert.ToString("SELECT * FROM ") & tableName) + " ORDER BY ") & columnName

                Using da = New SqlCeDataAdapter(sql, Connection())
                    da.Fill(DateFormatTable)
                End Using

                DateFormatTable.DefaultView.Sort = columnName & Convert.ToString(" asc")
                DateFormatTable.TableName = tableName

            Catch ex As Exception
                'ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub CreateTimeFormatTable()
            Try
                Dim columnName As String = "FormatString"
                Dim dcFormatString As New DataColumn(columnName, GetType(String))
                TimeFormatTable.Rows.Clear()
                Dim columns As DataColumnCollection = TimeFormatTable.Columns
                If columns.Contains(columnName) = False Then
                    TimeFormatTable.Columns.Add(dcFormatString)
                End If

                Dim tableName As String = "TimeFormat"
                Dim sql As String = Convert.ToString((Convert.ToString("SELECT * FROM ") & tableName) + " ORDER BY ") & columnName

                Using da = New SqlCeDataAdapter(sql, Connection())
                    da.Fill(TimeFormatTable)
                End Using

                TimeFormatTable.DefaultView.Sort = columnName & Convert.ToString(" asc")
                TimeFormatTable.TableName = tableName

            Catch ex As Exception
                'ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub CreateGraphDataTable()
            Try
                Dim columnName As String = "ORDR_NBR"
                Dim dcFormatString As New DataColumn(columnName, GetType(String))
                GraphDataTable.Rows.Clear()
                Dim columns As DataColumnCollection = GraphDataTable.Columns
                If columns.Contains(columnName) = False Then
                    GraphDataTable.Columns.Add(dcFormatString)
                End If

                Dim tableName As String = "GraphData"
                Dim sql As String = Convert.ToString((Convert.ToString("SELECT * FROM ") & tableName) + " ORDER BY ") & columnName

                Using da = New SqlCeDataAdapter(sql, Connection())
                    da.Fill(GraphDataTable)
                End Using

                GraphDataTable.DefaultView.Sort = columnName & Convert.ToString(" asc")

            Catch ex As Exception
                'ErrorHandler.DisplayMessage(ex)

            Finally

            End Try

        End Sub

        Public Shared Sub SetServerPath()
            Try
                Dim versionNumber As String = "" 'AssemblyInfo.versionFolderNumber TODO: FIX THIS
                'return Path.Combine(My.Settings.App_PathDeploy, "Application Files", AssemblyInfo.Product + versionNumber, dataFolder, AssemblyInfo.Product + ".sdf.deploy"); //for internal server
                Dim baseUrl As New Uri(My.Settings.App_PathDeploy)
                Dim relativeUrl As String = (Convert.ToString("Application Files/" + My.Application.Info.ProductName) & versionNumber) + "/" + dataFolder + "/"
                Dim combined As New Uri(baseUrl, relativeUrl)

                My.Settings.App_PathDeployData = combined.ToString()

            Catch ex As Exception
                'ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub SetUserPath()
            Dim userFilePath As String = String.Empty
            Try
                If System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed Then
                    Dim versionNumber As String = "" 'AssemblyInfo.versionFolderNumber TODO: FIX THIS
                    Dim localAppData As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                    'userFilePath = Path.Combine(localAppData, AssemblyInfo.Copyright.Replace(" ", "_"), AssemblyInfo.Product, dataFolder)
                    'if (!Directory.Exists(userFilePath)) Directory.CreateDirectory(userFilePath);
                    System.IO.Directory.CreateDirectory(userFilePath)
                    Dim info As New DirectoryInfo(userFilePath)
                    Dim security As DirectorySecurity = info.GetAccessControl()
                    security.AddAccessRule(New FileSystemAccessRule("Everyone", FileSystemRights.Modify, InheritanceFlags.ContainerInherit, PropagationFlags.None, AccessControlType.Allow))
                    security.AddAccessRule(New FileSystemAccessRule("Everyone", FileSystemRights.Modify, InheritanceFlags.ObjectInherit, PropagationFlags.None, AccessControlType.Allow))

                    info.SetAccessControl(security)
                Else
                    userFilePath = "C:\Users\tduguid\Source\Repos\ScriptHelp\VB\App_Data" 'System.IO.Path.Combine(AssemblyInfo.GetClickOnceLocation(), dataFolder)
                End If

                My.Settings.App_PathLocalData = userFilePath

            Catch ex As Exception
                'ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub InsertRecord(ByVal tbl As DataTable, ByVal text As String)
            Dim tableName As String = tbl.TableName.ToString()
            Dim columnName As String = tbl.Columns(0).ColumnName.ToString()
            Dim sql As String = "SELECT * FROM " & tableName
            If tbl.[Select](columnName & " = '" + text.Replace("'", "''") & "'").Length = 0 Then
                Dim dr As DialogResult = MessageBox.Show("Would you like to add '" & text & "' to the list?", "Add New Value", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                Select Case dr
                    Case DialogResult.Yes
                        tbl.Rows.Add(New Object() {text})
                        Dim cn As SqlCeConnection = New SqlCeConnection(Data.Connection())
                        Dim scb As SqlCeCommandBuilder = Nothing
                        Dim sda As SqlCeDataAdapter = New SqlCeDataAdapter(sql, cn)
                        sda.TableMappings.Add("Table", tableName)
                        scb = New SqlCeCommandBuilder(sda)
                        sda.Update(tbl)
                        Dim dcFormatString As DataColumn = New DataColumn(columnName, GetType(String))
                        tbl.Rows.Clear()
                        Dim columns As DataColumnCollection = tbl.Columns
                        If columns.Contains(columnName) = False Then
                            tbl.Columns.Add(dcFormatString)
                        End If

                        Using da = New SqlCeDataAdapter(sql, Connection())
                            da.Fill(tbl)
                        End Using

                        tbl.DefaultView.Sort = columnName & " asc"
                    Case DialogResult.No
                End Select
            End If
        End Sub

    End Class

End Namespace