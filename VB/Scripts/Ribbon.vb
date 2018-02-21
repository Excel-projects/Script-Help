Option Strict Off 'late binding on Excel object
Option Explicit On

Imports System.Diagnostics
Imports System.IO.Path
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Tools.Ribbon
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows
Imports System.Windows.Forms
Imports ScriptHelp.Scripts

Namespace Scripts

    <Runtime.InteropServices.ComVisible(True)>
    Public Class Ribbon
        Implements Office.IRibbonExtensibility
        Private ribbon As Office.IRibbonUI

        Public Shared ribbonref As Ribbon

        Public Shared mySettings As Settings
        Public Shared myTaskPaneSettings As Microsoft.Office.Tools.CustomTaskPane

        Public Shared myScript As Script
        Public Shared myTaskPaneScript As Microsoft.Office.Tools.CustomTaskPane

        Public Shared myTableData As TableData
        Public Shared myTaskPaneTableData As Microsoft.Office.Tools.CustomTaskPane

        Public Class AppVariables
            Private Sub New()
            End Sub

            Public Shared Property ScriptRange() As String
                Get
                    Return m_ScriptRange
                End Get
                Set
                    m_ScriptRange = Value
                End Set
            End Property
            Private Shared m_ScriptRange As String

            Public Shared Property FileType() As String
                Get
                    Return m_FileType
                End Get
                Set
                    m_FileType = Value
                End Set
            End Property
            Private Shared m_FileType As String

            Public Shared Property TableName() As String
                Get
                    Return m_TableName
                End Get
                Set
                    m_TableName = Value
                End Set
            End Property
            Private Shared m_TableName As String

            Public Shared Property FirstColumnName() As String
                Get
                    Return m_FirstColumnName
                End Get
                Set
                    m_FirstColumnName = Value
                End Set
            End Property
            Private Shared m_FirstColumnName As String

        End Class

#Region "  Ribbon Events  "
        Public Sub New()
        End Sub

        Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
            Return GetResourceText("ScriptHelp.Ribbon.xml")
        End Function

        Private Shared Function GetResourceText(ByVal resourceName As String) As String
            Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
            Dim resourceNames() As String = asm.GetManifestResourceNames()
            For i As Integer = 0 To resourceNames.Length - 1
                If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                    Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                        If resourceReader IsNot Nothing Then
                            Return resourceReader.ReadToEnd()
                        End If
                    End Using
                End If
            Next
            Return Nothing
        End Function

        Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
            Try
                Me.ribbon = ribbonUI
                ribbonref = Me
                'ribbonref = Me.ribbon
                'ThisAddIn.e_ribbon = ribbonUI
                'AssemblyInfo.SetAddRemoveProgramsIcon("ExcelAddin.ico")
                'AssemblyInfo.SetAssemblyFolderVersion()
                Data.SetServerPath()
                Data.SetUserPath()
                'ErrorHandler.SetLogPath()
                ErrorHandler.CreateLogRecord()

                'Dim destFilePath As String = Path.Combine(Properties.Settings.[Default].App_PathLocalData, AssemblyInfo.Product + ".sdf")
                'If Not (File.Exists(destFilePath)) Then
                ' Using client = New System.Net.WebClient()
                'client.DownloadFile(Properties.Settings.[Default].App_PathDeployData + AssemblyInfo.Product + ".sdf.deploy", Path.Combine(Properties.Settings.[Default].App_PathLocalData, AssemblyInfo.Product + ".sdf"))

                'End Using
                'End If

                Data.CreateTableAliasTable()
                Data.CreateDateFormatTable()
                Data.CreateTimeFormatTable()
                Data.CreateGraphDataTable()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Function GetButtonImage(ByVal control As Office.IRibbonControl) As System.Drawing.Bitmap
            Try
                Select Case control.Id
                    Case "btnScriptTypeDqlAppend", "btnScriptTypeDqlAppendLocked", "btnScriptTypeDqlCreate", "btnScriptTypeDqlTruncateAppend", "btnScriptTypeDqlUpdate", "btnScriptTypeDqlUpdateLocked"
                        Return My.Resources.Resources.ScriptTypeDql
                    Case "btnScriptTypeTSqlCreateTable", "btnScriptTypeTSqlInsertValues", "btnScriptTypeTSqlMergeValues", "btnScriptTypeTSqlSelectValues", "btnScriptTypeTSqlSelectUnion", "btnScriptTypeTSqlUpdateValues"
                        Return My.Resources.Resources.ScriptTypeTSql
                    Case "btnScriptTypePlSqlCreateTable", "btnScriptTypePlSqlInsertValues", "btnScriptTypePlSqlMergeValues", "btnScriptTypePlSqlSelectValues", "btnScriptTypePlSqlSelectUnion", "btnScriptTypePlSqlUpdateValues"
                        Return My.Resources.Resources.ScriptTypePlSql
                    Case "btnScriptTypeGithubTable"
                        Return My.Resources.Resources.ScriptTypeMarkdown
                    Case "btnScriptTypeHtmlTable", "btnScriptTypeXmlValues"
                        Return My.Resources.Resources.ScriptTypeMarkup
                    Case "btnProblemStepRecorder"
                        Return My.Resources.Resources.problem_steps_recorder
                    Case "btnSnippingTool"
                        Return My.Resources.Resources.snipping_tool
                    Case Else
                        Return Nothing
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return Nothing

            End Try

        End Function

        Public Function GetLabelText(ByVal control As Office.IRibbonControl) As String
            Try
                Select Case control.Id.ToString
                    Case Is = "tabScriptHelp"
                        If Application.ProductVersion.Substring(0, 2) = "15" Then
                            Return My.Application.Info.Title.ToUpper()
                        Else
                            Return My.Application.Info.Title
                        End If
                    Case Is = "txtCopyright"
                        Return "© " & My.Application.Info.Copyright.ToString
                    Case Is = "txtDescription", "btnDescription"
                        Dim version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                        Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & version
                    Case Is = "txtReleaseDate"
                        Return My.Settings.App_ReleaseDate.ToString("dd-MMM-yyyy hh:mm tt")
                    Case Is = "btnScriptTypeDqlAppend"
                        Return "DQL Append"
                    Case Is = "btnScriptTypeDqlAppendLocked"
                        Return "DQL Append/Locked"
                    Case Is = "btnScriptTypeDqlCreate"
                        Return "DQL Create"
                    Case Is = "btnScriptTypeDqlTruncateAppend"
                        Return "DQL Truncate/Append"
                    Case Is = "btnScriptTypeDqlUpdate"
                        Return "DQL Update"
                    Case Is = "btnScriptTypeDqlUpdateLocked"
                        Return "DQL Update/Locked"
                    Case Is = "btnScriptTypeGithubTable"
                        Return "Markdown Table"
                    Case Is = "btnScriptTypeHtmlTable"
                        Return "HTML Table"
                    Case Is = "btnScriptTypePlSqlCreateTable"
                        Return "PL/SQL Create Table"
                    Case Is = "btnScriptTypePlSqlInsertValues"
                        Return "PL/SQL Insert Values"
                    Case Is = "btnScriptTypePlSqlMergeValues"
                        Return "PL/SQL Merge Values"
                    Case Is = "btnScriptTypePlSqlSelectValues"
                        Return "PL/SQL Select Values"
                    Case Is = "btnScriptTypePlSqlSelectUnion"
                        Return "PL/SQL Select Union"
                    Case Is = "btnScriptTypePlSqlUpdateValues"
                        Return "PL/SQL Update Values"
                    Case Is = "btnScriptTypeTSqlCreateTable"
                        Return "T-SQL Create Table"
                    Case Is = "btnScriptTypeTSqlInsertValues"
                        Return "T-SQL Insert Values"
                    Case Is = "btnScriptTypeTSqlMergeValues"
                        Return "T-SQL Merge Values"
                    Case Is = "btnScriptTypeTSqlSelectValues"
                        Return "T-SQL Select Values"
                    Case Is = "btnScriptTypeTSqlSelectUnion"
                        Return "T-SQL Select Union"
                    Case Is = "btnScriptTypeTSqlUpdateValues"
                        Return "T-SQL Update Values"
                    Case Is = "btnScriptTypeXmlValues"
                        Return "XML Values"
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Function GetItemCount(control As Office.IRibbonControl) As Integer
            Try
                Select Case control.Id
                    Case "cboFormatDate"
                        Return Data.DateFormatTable.Rows.Count
                    Case "cboFormatTime"
                        Return Data.TimeFormatTable.Rows.Count
                    Case "cboTableAlias"
                        Return Data.TableAliasTable.Rows.Count
                    Case Else
                        Return 0
                End Select
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return 0
            End Try
        End Function

        Public Function GetItemLabel(control As Office.IRibbonControl, index As Integer) As String
            Try
                Select Case control.Id
                    Case "cboFormatDate"
                        Return UpdateDateFormatComboBoxSource(index)
                    Case "cboFormatTime"
                        Return UpdateTimeFormatComboBoxSource(index)
                    Case "cboTableAlias"
                        Return UpdateTableAliasComboBoxSource(index)
                    Case Else
                        Return String.Empty
                End Select
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty
            End Try
        End Function

        Public Sub GetSelectedItemID(ByVal Control As Office.IRibbonControl, ByRef itemID As Object)
            Try
                Select Case Control.Id.ToString
                    Case Is = "drpQueryType"
                        itemID = "UPDATE"
                    Case Else
                        itemID = String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                itemID = String.Empty

            End Try

        End Sub

        Public Function GetCount(ByVal Control As Office.IRibbonControl) As Integer
            Try
                Select Case Control.Id.ToString
                    Case Is = "drpQueryType"
                        Return 2
                    Case Else
                        Return 0
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return 0

            End Try

        End Function

        Public Function GetLabel(ByVal Control As Office.IRibbonControl, Index As Integer) As String
            Try
                Select Case Control.Id.ToString
                    Case Is = "drpQueryType"
                        Select Case Index
                            Case 0
                                Return "INSERT"
                            Case 1
                                Return "UPDATE"
                            Case Else
                                Return String.Empty
                        End Select
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Sub OnAction(ByVal Control As Office.IRibbonControl)
            Try
                Select Case Control.Id
                    Case "btnStart"
                    'OpenGraphData()
                    Case "btnCopyVisibleCells"
                        CopyVisibleCells()
                    Case "btnCleanData"
                        CleanData()
                    Case "btnZeroToNull"
                        ZeroStringToNull()
                    Case "btnFormatDateColumns"
                        FormatDateColumns()
                    Case "btnFormatTimeColumns"
                        FormatTimeColumns()
                    Case "btnClearInteriorColor"
                        ClearInteriorColor()
                    Case "btnSeparateValues"
                        SeparateValues()
                    Case "btnSettings"
                        OpenSettings()
                    Case "btnFileList"
                        CreateFileList()
                    Case "btnOpenReadMe"
                        OpenReadMe()
                    Case "btnOpenNewIssue"
                        OpenNewIssue()
                    Case "btnDownloadNewVersion"
                    'DownloadNewVersion();
                    Case "btnScriptTypeDqlAppend"
                        Formula.DqlAppend()
                    Case "btnScriptTypeDqlAppendLocked"
                        Formula.DqlAppendLocked()
                    Case "btnScriptTypeDqlCreate"
                        Formula.DqlCreate()
                    Case "btnScriptTypeDqlTruncateAppend"
                        Formula.DqlTruncateAppend()
                    Case "btnScriptTypeDqlUpdate"
                        Formula.DqlUpdate()
                    Case "btnScriptTypeDqlUpdateLocked"
                        Formula.DqlUpdateLocked()
                    Case "btnScriptTypeGithubTable"
                        Formula.GithubTable()
                    Case "btnScriptTypeHtmlTable"
                        Formula.HtmlTable()
                    Case "btnScriptTypePlSqlCreateTable"
                        Formula.PlSqlCreateTable()
                    Case "btnScriptTypePlSqlInsertValues"
                        Formula.PlSqlInsertValues()
                    Case "btnScriptTypePlSqlMergeValues"
                        Formula.PlSqlMergeValues()
                    Case "btnScriptTypePlSqlSelectValues"
                        Formula.PlSqlSelectValues()
                    Case "btnScriptTypePlSqlSelectUnion"
                        Formula.PlSqlSelectUnion()
                    Case "btnScriptTypePlSqlUpdateValues"
                        Formula.PlSqlUpdateValues()
                    Case "btnScriptTypeTSqlCreateTable"
                        Formula.TSqlCreateTable()
                    Case "btnScriptTypeTSqlInsertValues"
                        Formula.TSqlInsertValues()
                    Case "btnScriptTypeTSqlMergeValues"
                        Formula.TSqlMergeValues()
                    Case "btnScriptTypeTSqlSelectValues"
                        Formula.TSqlSelectValues()
                    Case "btnScriptTypeTSqlSelectUnion"
                        Formula.TSqlSelectUnion()
                    Case "btnScriptTypeTSqlUpdateValues"
                        Formula.TSqlUpdateValues()
                    Case "btnScriptTypeXmlValues"
                        Formula.XmlValues()
                    Case "btnFormatDate", "btnTableAlias", "btnFormatTime"
                        AppVariables.TableName = Control.Tag
                        OpenTableDataPane()
                    Case "btnSnippingTool"
                        OpenSnippingTool()
                    Case "btnProblemStepRecorder"
                        OpenProblemStepRecorder()
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Sub OnChange(control As Office.IRibbonControl, text As String)
            Try
                Select Case control.Id
                    Case "cboFormatDate"
                        My.Settings.Table_ColumnFormatDate = text
                        Data.InsertRecord(Data.DateFormatTable, text)
                        Exit Select
                    Case "cboFormatTime"
                        My.Settings.Table_ColumnFormatTime = text
                        Data.InsertRecord(Data.TimeFormatTable, text)
                        Exit Select
                    Case "cboTableAlias"
                        My.Settings.Table_ColumnTableAlias = text
                        Data.InsertRecord(Data.TableAliasTable, text)
                        Exit Select
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try

        End Sub

        Public Function GetText(control As Office.IRibbonControl) As String
            Try
                Select Case control.Id
                    Case "cboFormatDate"
                        Return My.Settings.Table_ColumnFormatDate
                    Case "cboFormatTime"
                        Return My.Settings.Table_ColumnFormatTime
                    Case "cboTableAlias"
                        Return My.Settings.Table_ColumnTableAlias
                    Case Else
                        Return String.Empty
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty
            End Try

        End Function

        Public Function GetVisible(ByVal Control As Office.IRibbonControl) As Boolean
            Try

                Select Case Control.Id
                    Case "grpClipboard"
                        Return My.Settings.Visible_grpClipboard
                    Case "grpAnnotation"
                        Return My.Settings.Visible_grpAnnotation
                    Case "btnScriptTypeTSqlCreateTable", "btnScriptTypeTSqlInsertValues", "btnScriptTypeTSqlMergeValues", "btnScriptTypeTSqlSelectValues", "btnScriptTypeTSqlSelectUnion", "btnScriptTypeTSqlUpdateValues"
                        Return My.Settings.Visible_mnuScriptType_TSQL
                    Case "btnScriptTypePlSqlCreateTable", "btnScriptTypePlSqlInsertValues", "btnScriptTypePlSqlMergeValues", "btnScriptTypePlSqlSelectValues", "btnScriptTypePlSqlSelectUnion", "btnScriptTypePlSqlUpdateValues"
                        Return My.Settings.Visible_mnuScriptType_PLSQL
                    Case "btnScriptTypeDqlAppend", "btnScriptTypeDqlAppendLocked", "btnScriptTypeDqlCreate", "btnScriptTypeDqlTruncateAppend", "btnScriptTypeDqlUpdate", "btnScriptTypeDqlUpdateLocked"
                        Return My.Settings.Visible_mnuScriptType_DQL
                    Case "btnScriptTypeGithubTable"
                        Return My.Settings.Visible_mnuScriptType_Markdown
                    Case "btnScriptTypeHtmlTable", "btnScriptTypeXmlValues"
                        Return My.Settings.Visible_mnuScriptType_Markup
                    Case Else
                        Return False
                End Select

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False

            End Try

        End Function

        Public Function GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
            Try
                Return True

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False

            End Try

        End Function

#End Region

#Region "  Ribbon Buttons  "

        Public Sub CopyVisibleCells()
            Dim visibleRange As Excel.Range = Nothing
            Try
                If ErrorHandler.IsEnabled(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                visibleRange = Globals.ThisAddIn.Application.Selection.SpecialCells(Excel.XlCellType.xlCellTypeVisible)
                visibleRange.Copy()

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                If visibleRange IsNot Nothing Then
                    'Marshal.ReleaseComObject(visibleRange)
                End If
            End Try

        End Sub

        Public Sub CleanData()
            Dim tbl As Excel.ListObject = Nothing
            Dim cell As Excel.Range = Nothing
            Dim usedRange As Excel.Range = Nothing
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                cell = Nothing
                Dim c As String = String.Empty
                Dim cc As String = String.Empty
                Dim cnt As Integer = 0
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                usedRange = tbl.Range
                Dim n As Integer = tbl.ListColumns.Count
                Dim m As Integer = tbl.ListRows.Count
                For i As Integer = 0 To m
                    ' by row
                    For j As Integer = 1 To n
                        ' by column
                        If usedRange(i + 1, j).Value2 IsNot Nothing Then
                            c = usedRange(i + 1, j).Value2.ToString()
                            ' can't convert null to string
                            If Globals.ThisAddIn.Application.WorksheetFunction.IsText(c) Then
                                cc = Globals.ThisAddIn.Application.WorksheetFunction.Clean(c.Trim())
                                If (cc <> c) Then
                                    cell = tbl.Range.Cells(i + 1, j)
                                    If Convert.ToBoolean(cell.HasFormula) = False Then
                                        cell.Value = cc
                                        cell.Interior.Color = My.Settings.Table_ColumnCleanedColour
                                        cnt = cnt + 1
                                    End If
                                End If
                                cell = tbl.Range.Cells(i + 1, j)
                                Dim qt As String = My.Settings.Table_ColumnScriptQuote
                                If cell.PrefixCharacter = qt Then
                                    ' show the leading apostrophe in the cell by doubling the value.
                                    cell.Value = (qt & qt) + cell.Value
                                    cell.Interior.Color = My.Settings.Table_ColumnCleanedColour
                                End If
                            End If
                        End If
                    Next
                Next
                MessageBox.Show("The number of cells cleaned: " + cnt.ToString(), "Cleaning has finished", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If cell IsNot Nothing Then
                    'Marshal.ReleaseComObject(cell)
                End If
                If usedRange IsNot Nothing Then
                    'Marshal.ReleaseComObject(usedRange)
                End If
            End Try

        End Sub

        Public Sub ZeroStringToNull()
            Dim tbl As Excel.ListObject = Nothing
            Dim cell As Excel.Range = Nothing
            Dim usedRange As Excel.Range = Nothing
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                cell = Nothing
                Dim cnt As Integer = 0
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                usedRange = tbl.Range
                Dim n As Integer = tbl.ListColumns.Count
                Dim m As Integer = tbl.ListRows.Count
                For i As Integer = 0 To m
                    ' by row
                    For j As Integer = 1 To n
                        ' by column
                        If usedRange(i + 1, j).Value2 Is Nothing Then
                            cell = tbl.Range.Cells(i + 1, j)
                            cell.Value = My.Settings.Table_ColumnScriptNull
                            cell.Interior.Color = My.Settings.Table_ColumnCleanedColour
                            cnt = cnt + 1
                        End If
                    Next
                Next
                MessageBox.Show("The number of cells converted to " + My.Settings.Table_ColumnScriptNull + ": " + cnt.ToString(), "Converting has finished", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If cell IsNot Nothing Then
                    'Marshal.ReleaseComObject(cell)
                End If
                If usedRange IsNot Nothing Then
                    'Marshal.ReleaseComObject(usedRange)
                End If
            End Try

        End Sub

        Public Sub FormatDateColumns()
            Dim tbl As Excel.ListObject = Nothing
            Dim cell As Excel.Range = Nothing
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                cell = Nothing
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                For Each col As Excel.ListColumn In tbl.ListColumns
                    cell = FirstNotNullCellInColumn(col.DataBodyRange)
                    If ((cell IsNot Nothing)) Then
                        If cell.NumberFormat.ToString() = My.Settings.Table_ColumnFormatTime Or ErrorHandler.IsDate(cell.Value) Then
                            col.DataBodyRange.NumberFormat = My.Settings.Table_ColumnFormatDate
                            col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter
                        End If
                    End If
                Next

            Catch ex As System.Runtime.InteropServices.COMException
                ErrorHandler.DisplayMessage(ex)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If cell IsNot Nothing Then
                    'Marshal.ReleaseComObject(cell)
                End If
            End Try

        End Sub

        Public Sub FormatTimeColumns()
            Dim tbl As Excel.ListObject = Nothing
            Dim cell As Excel.Range = Nothing
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                cell = Nothing
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                For Each col As Excel.ListColumn In tbl.ListColumns
                    cell = FirstNotNullCellInColumn(col.DataBodyRange)
                    If ((cell IsNot Nothing)) Then
                        If cell.NumberFormat.ToString() = My.Settings.Table_ColumnFormatTime Or ErrorHandler.IsDate(cell.Value) Then
                            col.DataBodyRange.NumberFormat = My.Settings.Table_ColumnFormatDate
                            col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter
                        End If
                    End If
                Next

            Catch ex As System.Runtime.InteropServices.COMException
                ErrorHandler.DisplayMessage(ex)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If cell IsNot Nothing Then
                    'Marshal.ReleaseComObject(cell)
                End If
            End Try

        End Sub

        Public Sub FormatAsTable()
            Dim range As Excel.Range = Nothing
            Dim tableName As String = My.Application.Info.Title + " " + DateTime.Now.ToString("yyyy-MM-ddThh:mm:ss:fffzzz")
            Dim tableStyle As String = My.Settings.Table_StyleName
            Try
                If ErrorHandler.IsValidListObject(False) = True Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                range = Globals.ThisAddIn.Application.Selection
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

                range.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name = tableName
                range.[Select]()
                range.Worksheet.ListObjects(tableName).TableStyle = tableStyle

                ribbon.ActivateTab("tabScriptHelp")

            Catch ex As System.Runtime.InteropServices.COMException
                ErrorHandler.DisplayMessage(ex)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If range IsNot Nothing Then
                    'Marshal.ReleaseComObject(range)
                End If
            End Try

        End Sub

        Public Sub ClearInteriorColor()
            Dim tbl As Excel.ListObject = Nothing
            Dim rng As Excel.Range = Nothing
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                tbl.DataBodyRange.Interior.ColorIndex = Excel.Constants.xlNone
                tbl.DataBodyRange.Font.ColorIndex = Excel.Constants.xlAutomatic
                rng = tbl.Range
                For i As Integer = 1 To rng.Columns.Count
                    If rng.Columns.EntireColumn(i).Hidden = False Then
                        DirectCast(rng.Cells(1, i), Excel.Range).Interior.ColorIndex = Excel.Constants.xlNone
                        DirectCast(rng.Cells(1, i), Excel.Range).HorizontalAlignment = Excel.Constants.xlCenter
                        DirectCast(rng.Cells(1, i), Excel.Range).VerticalAlignment = Excel.Constants.xlCenter
                    End If
                Next

            Catch ex As System.Runtime.InteropServices.COMException
                ErrorHandler.DisplayMessage(ex)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If rng IsNot Nothing Then
                    'Marshal.ReleaseComObject(rng)
                End If
            End Try

        End Sub

        Public Sub SeparateValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim cell As Excel.Range = Nothing
            Try
                If ErrorHandler.IsAvailable(True) = False Then
                    Return
                End If
                ErrorHandler.CreateLogRecord()
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                cell = Globals.ThisAddIn.Application.ActiveCell
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim m As Integer = tbl.ListRows.Count
                Dim a As Integer = m
                Dim columnIndex As Integer = cell.Column

                For i As Integer = 1 To m + 1
                    ' by row
                    Dim cellValue As String = tbl.Range.Cells(i, columnIndex).Value2.ToString()
                    If String.IsNullOrEmpty(cellValue) = False Then
                        Dim metadata As String() = cellValue.Split(My.Settings.Table_ColumnSeparateValuesDelimiter)
                        Dim countValues As Integer = metadata.Length - 1
                        If countValues > 0 Then
                            'if the column value has multiple values then create a new row per value
                            For j As Integer = 1 To countValues
                                ' by value 
                                tbl.ListRows.Add(i)
                                tbl.Range.Rows(i + 1).Value = tbl.Range.Rows(i).Value
                                ' get the next value in the string
                                tbl.Range.Cells(i + 1, columnIndex).Value2 = metadata(j - 1).Trim()
                            Next
                            tbl.Range.Cells(i, columnIndex).Value2 = metadata(countValues).Trim()
                            ' reset the first row value
                            m += countValues
                            'reset the total rows
                            'reset the current row
                            i += countValues
                        End If

                    End If
                Next
                MessageBox.Show("The number of row(s) added is " + (m - a).ToString(), "Finished Separating Values", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If cell IsNot Nothing Then
                    'Marshal.ReleaseComObject(cell)
                End If
            End Try

        End Sub

        Public Sub CreateFileList()
            Dim filePath As String = "" 'My.Settings.Option_PathFileListing
            Try
                ErrorHandler.CreateLogRecord()
                Dim msgDialogResult As DialogResult = DialogResult.None
                Dim dlg As New FolderBrowserDialog()
            'If My.Settings.Option_PathFileListingSelect = True Then
            dlg.RootFolder = Environment.SpecialFolder.MyComputer
                    dlg.SelectedPath = filePath
                    msgDialogResult = dlg.ShowDialog()
                    filePath = dlg.SelectedPath
            'End If
            'If msgDialogResult = DialogResult.OK Or My.Settings.Option_PathFileListingSelect = False Then
            filePath += "\"
                    Dim scriptCommands As String = String.Empty
                    Dim currentDate As String = DateTime.Now.ToString("dd.MMM.yyyy_hh.mm.tt")
                    Dim batchFileName As String = (Convert.ToString(filePath & Convert.ToString("FileListing_")) & currentDate) + "_" + Environment.UserName + ".bat"
                    scriptCommands = "echo off" + Environment.NewLine
                    scriptCommands += "cd %1" + Environment.NewLine
                    scriptCommands += (Convert.ToString((Convert.ToString((Convert.ToString("dir """) & filePath) + """ /s /a-h /b /-p /o:gen >""") & filePath) + "FileListing_") & currentDate) + "_" + Environment.UserName + ".csv""" + Environment.NewLine
                    scriptCommands += (Convert.ToString((Convert.ToString("""") & filePath) + "FileListing_") & currentDate) + "_" + Environment.UserName + ".csv""" + Environment.NewLine
                    scriptCommands += "cd .. " + Environment.NewLine
                    scriptCommands += "echo on" + Environment.NewLine
                    System.IO.File.WriteAllText(batchFileName, scriptCommands)
            'AssemblyInfo.OpenFile(batchFileName)
            'End If

            Catch generatedExceptionName As System.UnauthorizedAccessException
                MessageBox.Show(Convert.ToString("You don't have access to this folder, bro!" + Environment.NewLine + Environment.NewLine) & filePath, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try

        End Sub

        Public Sub OpenReadMe()
            ErrorHandler.CreateLogRecord()
            System.Diagnostics.Process.Start(My.Settings.App_PathReadMe)
        End Sub

        Public Sub OpenNewIssue()
            ErrorHandler.CreateLogRecord()
            System.Diagnostics.Process.Start(My.Settings.App_PathNewIssue)

        End Sub

        Public Sub OpenSettings()
            Try
                If myTaskPaneSettings IsNot Nothing Then
                    If myTaskPaneSettings.Visible = True Then
                        myTaskPaneSettings.Visible = False
                    Else
                        myTaskPaneSettings.Visible = True
                    End If
                Else
                    mySettings = New Settings()
                    myTaskPaneSettings = Globals.ThisAddIn.CustomTaskPanes.Add(mySettings, "Settings for " + My.Application.Info.Title)
                    myTaskPaneSettings.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                    myTaskPaneSettings.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                    myTaskPaneSettings.Width = 675
                    myTaskPaneSettings.Visible = True

                End If
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try
        End Sub

        Public Shared Sub OpenScriptPane()
            Try
                If myTaskPaneScript IsNot Nothing Then
                    myTaskPaneScript.Dispose()
                    myScript.Dispose()
                End If
                myScript = New Script()
                myTaskPaneScript = Globals.ThisAddIn.CustomTaskPanes.Add(myScript, "Script for " + My.Application.Info.Title)
                myTaskPaneScript.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                myTaskPaneScript.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                myTaskPaneScript.Width = 675

                myTaskPaneScript.Visible = True

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try

        End Sub

        Public Sub OpenTableDataPane()
            Try
                If myTaskPaneTableData IsNot Nothing Then
                    myTaskPaneTableData.Dispose()
                    myTableData.Dispose()
                End If
                myTableData = New TableData()
                myTaskPaneTableData = Globals.ThisAddIn.CustomTaskPanes.Add(myTableData, "List of " + AppVariables.TableName + " for " + My.Application.Info.Title)
                myTaskPaneTableData.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight
                myTaskPaneTableData.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                myTaskPaneTableData.Width = 300

                myTaskPaneTableData.Visible = True

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

#End Region

#Region "  Subroutines  "

        Friend Shared Function ApplyTextQuotes(ByVal col As Excel.ListColumn) As String
            Try
                If (GetSqlDataType(col) <> My.Settings.Column_TypeNumeric) Then
                    Return My.Settings.Table_ColumnScriptQuote
                Else
                    Return String.Empty
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Public Shared Function ConcatenateColumnNames(rng As Excel.Range, Optional tableAliasName As String = "", Optional prefixChar As String = "", Optional suffixChar As String = "", Optional selectionChar As String = ", ") As String
            Try
                Dim columnNames As String = String.Empty
                If tableAliasName <> String.Empty Then
                    tableAliasName = tableAliasName & Convert.ToString(".")
                End If
                For i As Integer = 1 To rng.Columns.Count - 1
                    If rng.Columns.EntireColumn(i).Hidden = False Then
                        columnNames = Convert.ToString((Convert.ToString(Convert.ToString(columnNames & selectionChar) & tableAliasName) & prefixChar) + DirectCast(rng.Cells(1, i), Excel.Range).Value2) & suffixChar
                    End If
                Next
                If columnNames.Substring(0, selectionChar.Length).Contains(selectionChar) AndAlso selectionChar.Length > 0 Then
                    columnNames = columnNames.Substring(2, columnNames.Length - 2)
                End If
                Return columnNames

            Catch generatedExceptionName As Exception
                Return String.Empty
            End Try

        End Function

        Public Shared Function ConcatenateColumnNamesJoin(rng As Excel.Range, tableAliasNameTarget As String, tableAliasNameSource As String) As String
            Try
                Dim columnNames As String = String.Empty
                For i As Integer = 1 To rng.Columns.Count - 1
                    If rng.Columns.EntireColumn(i).Hidden = False Then
                        columnNames = (Convert.ToString((Convert.ToString(columnNames & Convert.ToString(", ")) & tableAliasNameTarget) + ".[" + DirectCast(rng.Cells(1, i), Excel.Range).Value2 + "] = ") & tableAliasNameSource) + ".[" + DirectCast(rng.Cells(1, i), Excel.Range).Value2 + "]" + Environment.NewLine
                    End If
                Next
                columnNames = columnNames.Substring(2, columnNames.Length - 2)
                Return columnNames

            Catch generatedExceptionName As Exception
                Return String.Empty
            End Try

        End Function

        Public Shared Function GetColumnFormat(ByVal col As Excel.ListColumn) As String
            Try
                Dim fmt As String = String.Empty
                Dim nFmt As String = String.Empty

                Select Case GetSqlDataType(col)
                    Case My.Settings.Column_TypeDate
                        fmt = My.Settings.Table_ColumnFormatDate

                    Case My.Settings.Column_TypeNumeric
                        ' we will use the column formatting if some is applied
                        If Not IsNothing(col.DataBodyRange.NumberFormat) Then
                            'If Not IsNull(col.DataBodyRange.NumberFormat) Then
                            nFmt = col.DataBodyRange.NumberFormat.ToString
                            If Not (nFmt = "General") Then
                                fmt = nFmt
                            End If
                        End If

                End Select
                Return Formatted(col, fmt)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Friend Shared Function Formatted(ByVal col As Excel.ListColumn, fmt As String) As String
            Try
                Formatted = "[" & col.Name & "]"
                If (fmt = "") Then
                    Exit Try
                End If
                Return "TEXT(" & Formatted & ",""" & fmt & """)"

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty

            End Try

        End Function

        Friend Shared Function GetSqlDataType(ByVal col As Excel.ListColumn) As Integer
            Try
                ' Determine the likely SQL type of the column
                ' default to text
                Return My.Settings.Column_TypeText
                Dim rowCnt As Integer = col.DataBodyRange.Rows.Count
                Dim numCnt As Double = 0
                Dim notNullCnt As Double = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(col.DataBodyRange, "<>" & My.Settings.Table_ColumnScriptNull)

                ' If all values are nulls then assume text
                If (notNullCnt = 0) Then
                    Return 0
                End If

                numCnt = Globals.ThisAddIn.Application.WorksheetFunction.Count(col.DataBodyRange)
                ' if no numbers then assume text
                If (numCnt = 0) Then
                    Return 0
                End If

                ' if a mix of numbers and not numbers then assume text
                If (numCnt <> notNullCnt) Then
                    Return 0
                End If

                ' NOTE: next test relies consistent formatting on numerics in a column
                ' so we only have to test the first cell
                If IsDate(FirstNotNullCellInColumn(col.DataBodyRange)) Or col.DataBodyRange.NumberFormat.ToString = My.Settings.Table_ColumnFormatDate Then
                    Return My.Settings.Column_TypeDate
                Else
                    Return My.Settings.Column_TypeNumeric
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return 0

            End Try

        End Function

        Friend Shared Function GetItem(ByVal col As Excel.ListColumns, key As String) As Excel.ListColumn
            Try
                Return col(key)

            Catch ex As Exception
                ' ErrorMsg(ex)            
                Return Nothing

            End Try

        End Function

        Public Sub InvalidateRibbon()
            ribbon.Invalidate()
        End Sub

        Friend Shared Function ListColumn(ByVal cell As Excel.Range) As Excel.ListColumn
            Try
                ' Get the list column for a cell
                'Dim c As Excel.Range = cell.Cells(1, 1) ' use top left most cell
                Dim c As Excel.Range = CType(cell.Cells(1, 1), Global.Microsoft.Office.Interop.Excel.Range) ' use top left most cell
                Dim tbl As Excel.ListObject = c.ListObject
                If (tbl Is Nothing) Then
                    Return Nothing
                    Exit Try
                End If
                Dim colNum As Integer = c.Column - tbl.DataBodyRange.Cells(1, 1).Column + 1
                Return tbl.ListColumns(colNum)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return Nothing
                Exit Try

            End Try

        End Function

        Friend Shared Function FirstNotNullCellInColumn(ByVal rng As Excel.Range) As Excel.Range
            ' TODO: find a way to do this without looping.
            ' NOTE: SpecialCells is unreliable when called from VBA UDFs (Odd ??!)        
            Try
                If (rng Is Nothing) Then
                    Return Nothing
                End If
                Dim cell As Excel.Range

                For Each cell In rng
                    If (cell.Value IsNot Nothing) Then
                        If (cell.Value.ToString <> My.Settings.Table_ColumnScriptNull) Then
                            Return cell
                        End If
                    End If
                Next
                Return Nothing

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return Nothing

            End Try

        End Function

        Public Sub OpenFile(ByVal FilePath As String)
            Try
                Dim pStart As New System.Diagnostics.Process
                If FilePath = String.Empty Then Exit Try
                pStart.StartInfo.FileName = FilePath
                pStart.Start()

            Catch ex As System.ComponentModel.Win32Exception
                'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & FilePath, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Try

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Function UpdateTableAliasComboBoxSource(itemIndex As Integer) As String
            Try
                Return Data.TableAliasTable.Rows(itemIndex)("TableName").ToString()
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty
            End Try

        End Function

        Public Function UpdateDateFormatComboBoxSource(itemIndex As Integer) As String
            Try
                Return Data.DateFormatTable.Rows(itemIndex)("FormatString").ToString()
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty
            End Try
        End Function

        Public Function UpdateTimeFormatComboBoxSource(itemIndex As Integer) As String
            Try
                Return Data.TimeFormatTable.Rows(itemIndex)("FormatString").ToString()
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return String.Empty
            End Try
        End Function

        Public Sub OpenSnippingTool()
            Dim filePath As String
            Try
                If System.Environment.Is64BitOperatingSystem Then
                    filePath = "C:\Windows\sysnative\SnippingTool.exe"
                Else
                    filePath = "C:\Windows\system32\SnippingTool.exe"
                End If

                System.Diagnostics.Process.Start(filePath)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try
        End Sub

        Public Sub OpenProblemStepRecorder()
            Dim filePath As String = "C:\Windows\System32\psr.exe"
            Try
                System.Diagnostics.Process.Start(filePath)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try
        End Sub

#End Region

    End Class

End Namespace