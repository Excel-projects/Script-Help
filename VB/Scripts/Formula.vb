Option Strict Off
Option Explicit On

Imports System.Windows.Forms

Namespace Scripts
    Public Class Formula

        Public Shared Sub DqlAppend()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty
                Dim afterWhere As Boolean = False

                For Each col As Excel.ListColumn In tbl.ListColumns
                    Select Case col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c})
                        Case -1
                            Exit Select
                        Case Else
                            MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                    End Select
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim columnName As String = col.Name
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)
                        If [String].Compare(firstPart.Trim(), "WHERE", True) = 0 Then
                            valuePrefix = """" + col.Name + " = "" & "
                            afterWhere = True
                        Else
                            If afterWhere = True Then
                                valuePrefix = """AND " + col.Name + " = "" & "
                            Else
                                valuePrefix = """APPEND " + col.Name + " = "" & "
                            End If
                        End If
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next
                'replace NULL values with DQL format
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "', '" + My.Settings.Table_ColumnFormatDate + "'"", ""'nulldate'"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "'"", ""nullstring"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", """ + My.Settings.Table_ColumnScriptNull + """, ""nullint"")"

                Dim whereCheck As String = String.Empty
                If afterWhere = False Then
                    whereCheck = " & "" WHERE "" "
                End If
                formula = (Convert.ToString(Convert.ToString((Convert.ToString("=""UPDATE ") & tableAlias) + " objects "" & ") & formula) & whereCheck) + " & CHAR(10) & ""GO """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "DQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)

            End Try

        End Sub

        Public Shared Sub DqlAppendLocked()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty
                Dim afterWhere As Boolean = False
                Dim whereClause As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    Select Case col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c})
                        Case -1
                            Exit Select
                        Case Else
                            MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                    End Select
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim columnName As String = col.Name
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)
                        If [String].Compare(firstPart.Trim(), "WHERE", True) = 0 Then
                            valuePrefix = """" + col.Name + " = "" & "
                            afterWhere = True
                        Else
                            If afterWhere = True Then
                                valuePrefix = """AND " + col.Name + " = "" & "
                            Else
                                valuePrefix = """APPEND " + col.Name + " = "" & "
                            End If
                        End If
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            If afterWhere = True Then

                                whereClause += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                            End If
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            If afterWhere = True Then
                                whereClause += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                            End If
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next
                'replace NULL values with DQL format
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "', '" + My.Settings.Table_ColumnFormatDate + "'"", ""'nulldate'"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "'"", ""nullstring"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", """ + My.Settings.Table_ColumnScriptNull + """, ""nullint"")"

                Dim whereCheck As String = String.Empty
                If afterWhere = False Then
                    MessageBox.Show("This update statement must have a WHERE clause.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    whereCheck = " & "" WHERE "" "
                End If
                Dim formulaUnlock As String = (Convert.ToString((Convert.ToString("""UPDATE ") & tableAlias) + "(all) objects SET r_immutable_flag = 0 "" & ") & whereClause) + " & CHAR(10) & ""GO ""  & CHAR(10) & "
                Dim formulaLock As String = (Convert.ToString((Convert.ToString("""UPDATE ") & tableAlias) + "(all) objects SET r_immutable_flag = 1 "" & ") & whereClause) + " & CHAR(10) & ""GO """
                formula = Convert.ToString((Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("=") & formulaUnlock) + """UPDATE ") & tableAlias) + "(all) objects "" & ") & formula) & whereCheck) + " & CHAR(10) & ""GO "" & CHAR(10) & ") & formulaLock
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "DQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub DqlCreate()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & CHAR(10) & "","" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        Dim valuePrefix As String = String.Empty
                        valuePrefix = """ SET " + col.Name + " = "" & "
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next
                'replace NULL values with DQL format
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "', '" + My.Settings.Table_ColumnFormatDate + "'"", ""'nulldate'"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "'"", ""nullstring"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", """ + My.Settings.Table_ColumnScriptNull + """, ""nullint"")"

                formula = (Convert.ToString((Convert.ToString("=""CREATE ") & tableAlias) + " objects "" & CHAR(10) & ") & formula) + " & ""; "" & CHAR(10) & ""GO """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "DQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub DqlTruncateAppend()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty
                Dim afterWhere As Boolean = False

                For Each col As Excel.ListColumn In tbl.ListColumns
                    Select Case col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c})
                        Case -1
                            Exit Select
                        Case Else
                            MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                    End Select
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim columnName As String = col.Name
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)
                        If [String].Compare(firstPart.Trim(), "WHERE", True) = 0 Then
                            valuePrefix = """" + col.Name + " = "" & "
                            afterWhere = True
                        Else
                            If afterWhere = True Then
                                valuePrefix = """AND " + col.Name + " = "" & "
                            Else
                                valuePrefix = """TRUNCATE " + col.Name + ", APPEND " + col.Name + " = "" & "
                            End If
                        End If
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next
                'replace NULL values with DQL format
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "', '" + My.Settings.Table_ColumnFormatDate + "'"", ""'nulldate'"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "'"", ""nullstring"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", """ + My.Settings.Table_ColumnScriptNull + """, ""nullint"")"

                Dim whereCheck As String = String.Empty
                If afterWhere = False Then
                    MessageBox.Show("This update statement must have a WHERE clause.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    whereCheck = " & "" WHERE "" "
                End If
                formula = (Convert.ToString(Convert.ToString((Convert.ToString("=""UPDATE ") & tableAlias) + " objects "" & ") & formula) & whereCheck) + " & CHAR(10) & ""GO """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "DQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub DqlUpdate()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty
                Dim afterWhere As Boolean = False

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim columnName As String = col.Name
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)
                        If [String].Compare(firstPart.Trim(), "WHERE", True) = 0 Then
                            valuePrefix = """" + col.Name + " = "" & "
                            afterWhere = True
                        Else
                            If afterWhere = True Then
                                valuePrefix = """AND " + col.Name + " = "" & "
                            Else
                                valuePrefix = """SET " + col.Name + " = "" & "
                            End If
                        End If
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next
                'replace NULL values with DQL format
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "', '" + My.Settings.Table_ColumnFormatDate + "'"", ""'nulldate'"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "'"", ""nullstring"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", """ + My.Settings.Table_ColumnScriptNull + """, ""nullint"")"

                Dim whereCheck As String = String.Empty
                If afterWhere = False Then
                    whereCheck = " & "" WHERE "" "
                End If
                formula = (Convert.ToString(Convert.ToString((Convert.ToString("=""UPDATE ") & tableAlias) + " objects "" & ") & formula) & whereCheck) + " & CHAR(10) & ""GO """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "DQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub DqlUpdateLocked()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty
                Dim afterWhere As Boolean = False
                Dim whereClause As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    Select Case col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c})
                        Case -1
                            Exit Select
                        Case Else
                            MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Return
                    End Select
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim columnName As String = col.Name
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)
                        If [String].Compare(firstPart.Trim(), "WHERE", True) = 0 Then
                            valuePrefix = """" + col.Name + " = "" & "
                            afterWhere = True
                        Else
                            If afterWhere = True Then
                                valuePrefix = """AND " + col.Name + " = "" & "
                            Else
                                valuePrefix = """SET " + col.Name + " = "" & "
                            End If
                        End If
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            If afterWhere = True Then

                                whereClause += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                            End If
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            If afterWhere = True Then
                                whereClause += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                            End If
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next
                'replace NULL values with DQL format
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "', '" + My.Settings.Table_ColumnFormatDate + "'"", ""'nulldate'"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", ""'" + My.Settings.Table_ColumnScriptNull + "'"", ""nullstring"")"
                formula = (Convert.ToString("SUBSTITUTE(") & formula) + ", """ + My.Settings.Table_ColumnScriptNull + """, ""nullint"")"

                Dim whereCheck As String = String.Empty
                If afterWhere = False Then
                    MessageBox.Show("This update statement must have a WHERE clause.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    whereCheck = " & "" WHERE "" "
                End If
                Dim formulaUnlock As String = (Convert.ToString((Convert.ToString("""UPDATE ") & tableAlias) + "(all) objects SET r_immutable_flag = 0 "" & ") & whereClause) + " & CHAR(10) & ""GO ""  & CHAR(10) & "
                Dim formulaLock As String = (Convert.ToString((Convert.ToString("""UPDATE ") & tableAlias) + "(all) objects SET r_immutable_flag = 1 "" & ") & whereClause) + " & CHAR(10) & ""GO """
                formula = Convert.ToString((Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("=") & formulaUnlock) + """UPDATE ") & tableAlias) + "(all) objects "" & ") & formula) & whereCheck) + " & CHAR(10) & ""GO "" & CHAR(10) & ") & formulaLock
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "DQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub GithubTable()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias

                Dim sqlColName As String = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                sqlCol = tbl.ListColumns.Add()
                sqlCol.Name = lastColumnName
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count

                sqlCol.DataBodyRange.NumberFormat = "General"

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & ""|"" & ")
                        End If
                        formula += Ribbon.GetColumnFormat(col).ToString()
                    End If
                Next
                formula = (Convert.ToString("=""" + "|" + """ & ") & formula) + " & ""|"""
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "TXT"
                    Dim headerColumn As String = Ribbon.ConcatenateColumnNames(tbl.Range, String.Empty, "|", "") + "|" + Environment.NewLine
                    Dim headerSeparator As String = "|:" + New [String]("-"c, 10)
                    Dim headerLine As String = New System.Text.StringBuilder(headerSeparator.Length * lastColumnIndex).Insert(0, headerSeparator, lastColumnIndex).ToString().Substring(0, ((headerSeparator.Length * lastColumnIndex) - (headerSeparator.Length - 1))) + Environment.NewLine
                    Ribbon.AppVariables.ScriptRange = Convert.ToString(headerColumn & headerLine) & DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub HtmlTable()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = String.Empty
                        ' Ribbon.ApplyTextQuotes(col);
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim valueSuffix As String = String.Empty
                        Dim columnName As String = col.Name.ToLower()
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)

                        valuePrefix = " CHAR(10) & ""<td>"" & "
                        valueSuffix = " & ""</td>"" "
                        formula += Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """") & valueSuffix
                    End If
                Next

                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                formula = (Convert.ToString("=""<tr> "" & ") & formula) + " & CHAR(10) & ""</tr> """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "XML"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                    Ribbon.AppVariables.ScriptRange = "<table>" + Environment.NewLine + "<tr>" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "", "<th>", "</th>", Environment.NewLine) + Environment.NewLine + "</tr>" + Environment.NewLine + Ribbon.AppVariables.ScriptRange + "</table>"

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub PlSqlCreateTable()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim insertPrefix As String = (Convert.ToString("INSERT INTO ") & tableAlias) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES("
                formula = (Convert.ToString((Convert.ToString("=""") & insertPrefix) + """ & ") & formula) + " & "");"""
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Dim createTable As String = (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString("IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'") & tableAlias) + "') AND type in (N'U'))" + Environment.NewLine + "DROP TABLE ") & tableAlias) + Environment.NewLine + "; " + Environment.NewLine + "CREATE TABLE ") & tableAlias) + " (") & tableAlias) + "_ID [int] PRIMARY KEY IDENTITY(1,1) NOT NULL, " + Ribbon.ConcatenateColumnNames(tbl.Range, "", Environment.NewLine + "[", "] [varchar](max) NULL") + Environment.NewLine + ");" + Environment.NewLine

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = createTable & DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub PlSqlInsertValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim insertPrefix As String = (Convert.ToString("INSERT INTO ") & tableAlias) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES("
                formula = (Convert.ToString((Convert.ToString("=""") & insertPrefix) + """ & ") & formula) + " & "");"""
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName

                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)

                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try

            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub PlSqlMergeValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim tableAliasTemp As String = tableAlias & Convert.ToString("_source")
                Dim sqlColName As String = String.Empty

                sqlColName = (Convert.ToString("SELECT ") & tableAliasTemp) + ".*" + " FROM (VALUES"

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                ' Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                        'DO NOTHING - because the column is hidden or the last column with the sql script
                        'if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim firstRowNbr As Integer = tbl.Range(1, 1).Row + 1
                ' must use the offset for the first row number
                formula = (Convert.ToString("=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, "" "", "","") & " + """ ( "" & ") & formula) + " & "")"""
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                tbl.ShowTotals = True
                Dim totalsColumnValue As String = (Convert.ToString(") ") & tableAliasTemp) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") "
                tbl.TotalsRowRange(lastColumnIndex).Value2 = totalsColumnValue
                ' totals row has a maximum limit of 32,767 characters
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                    Ribbon.AppVariables.ScriptRange = (Convert.ToString((Convert.ToString((Convert.ToString("SET XACT_ABORT ON" + Environment.NewLine + "BEGIN TRANSACTION;" + Environment.NewLine + Environment.NewLine + ";WITH " + Environment.NewLine) & tableAliasTemp) + Environment.NewLine + "AS " + Environment.NewLine + "(" + Environment.NewLine + Ribbon.AppVariables.ScriptRange + ") " + Environment.NewLine + "MERGE ") & tableAlias) + " AS T" + Environment.NewLine + "USING ") & tableAliasTemp) + " AS S" + Environment.NewLine + "ON " + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "WHEN NOT MATCHED BY TARGET" + Environment.NewLine + "THEN INSERT" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "VALUES" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "S", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "WHEN MATCHED" + Environment.NewLine + "THEN UPDATE SET" + Environment.NewLine + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "--WHEN NOT MATCHED BY SOURCE AND 'ADD WHERE CLAUSE HERE'" + Environment.NewLine + "--THEN DELETE" + Environment.NewLine + "OUTPUT $action, inserted.*, deleted.*;" + Environment.NewLine + Environment.NewLine + "ROLLBACK TRANSACTION;" + Environment.NewLine + "--COMMIT TRANSACTION;" + Environment.NewLine + "GO"
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub PlSqlSelectValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty

                sqlColName = (Convert.ToString("SELECT ") & tableAlias) + ".*" + " FROM (VALUES"

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                ' Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                        'DO NOTHING - because the column is hidden or the last column with the sql script
                        'if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim firstRowNbr As Integer = tbl.Range(1, 1).Row + 1
                ' must use the offset for the first row number
                formula = (Convert.ToString("=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, "" "", "","") & " + """ ( "" & ") & formula) + " & "")"""
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                tbl.ShowTotals = True
                Dim totalsColumnValue As String = (Convert.ToString(") ") & tableAlias) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") "
                tbl.TotalsRowRange(lastColumnIndex).Value2 = totalsColumnValue
                ' totals row has a maximum limit of 32,767 characters
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub PlSqlSelectUnion()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        Dim dqt As String = """"""
                        Dim valuePlSuffix As String = (Convert.ToString((Convert.ToString("& "" AS ") & dqt) + col.Name) & dqt) + " """
                        formula += Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """") & valuePlSuffix
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim firstRowNbr As Integer = tbl.Range(1, 1).Row + 1
                ' must use the offset for the first row number
                formula = (Convert.ToString("=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, """", ""UNION "") & " + """SELECT "" & ") & formula) + " & "" FROM DUAL"""
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub PlSqlUpdateValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty
                Dim afterWhere As Boolean = False

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim columnName As String = col.Name
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)
                        If [String].Compare(firstPart.Trim(), "WHERE", True) = 0 Then
                            Dim firstWhereColumn As String = col.Name
                            firstWhereColumn = System.Text.RegularExpressions.Regex.Replace(firstWhereColumn, "where ", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                            firstWhereColumn = firstWhereColumn.Trim()
                            valuePrefix = (Convert.ToString("""WHERE ") & firstWhereColumn) + " = "" & "
                            afterWhere = True
                        Else
                            If afterWhere = True Then
                                valuePrefix = """AND " + col.Name + " = "" & "
                            Else
                                Dim useComma As String = String.Empty
                                If columnCount <> 1 Then
                                    useComma = ","
                                End If
                                valuePrefix = (Convert.ToString("""") & useComma) + " " + col.Name + " = "" & "
                            End If
                        End If
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next

                Dim whereCheck As String = String.Empty
                If afterWhere = False Then
                    whereCheck = " & "" WHERE "" "
                End If
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                formula = (Convert.ToString(Convert.ToString((Convert.ToString("=""UPDATE ") & tableAlias) + " SET "" & ") & formula) & whereCheck) + " & CHAR(10) & ""GO """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub TSqlCreateTable()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim insertPrefix As String = (Convert.ToString("INSERT INTO ") & tableAlias) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES("
                formula = (Convert.ToString((Convert.ToString("=""") & insertPrefix) + """ & ") & formula) + " & "");"""
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Dim createTable As String = (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString("IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'") & tableAlias) + "') AND type in (N'U'))" + Environment.NewLine + "DROP TABLE ") & tableAlias) + Environment.NewLine + "; " + Environment.NewLine + "CREATE TABLE ") & tableAlias) + " (") & tableAlias) + "_ID [int] PRIMARY KEY IDENTITY(1,1) NOT NULL, " + Ribbon.ConcatenateColumnNames(tbl.Range, "", Environment.NewLine + "[", "] [varchar](max) NULL") + Environment.NewLine + ");" + Environment.NewLine
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = createTable & DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub TSqlInsertValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim insertPrefix As String = (Convert.ToString("INSERT INTO ") & tableAlias) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range) + ") VALUES("
                formula = (Convert.ToString((Convert.ToString("=""") & insertPrefix) + """ & ") & formula) + " & "");"""
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub TSqlMergeValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try

                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim tableAliasTemp As String = tableAlias & Convert.ToString("_source")
                Dim sqlColName As String = String.Empty

                sqlColName = (Convert.ToString("SELECT ") & tableAliasTemp) + ".*" + " FROM (VALUES"

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                ' Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                        'DO NOTHING - because the column is hidden or the last column with the sql script
                        'if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim firstRowNbr As Integer = tbl.Range(1, 1).Row + 1
                ' must use the offset for the first row number
                formula = (Convert.ToString("=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, "" "", "","") & " + """ ( "" & ") & formula) + " & "")"""
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                tbl.ShowTotals = True
                Dim totalsColumnValue As String = (Convert.ToString(") ") & tableAliasTemp) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") "
                tbl.TotalsRowRange(lastColumnIndex).Value2 = totalsColumnValue
                ' totals row has a maximum limit of 32,767 characters
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                    Ribbon.AppVariables.ScriptRange = (Convert.ToString((Convert.ToString((Convert.ToString("SET XACT_ABORT ON" + Environment.NewLine + "BEGIN TRANSACTION;" + Environment.NewLine + Environment.NewLine + ";WITH " + Environment.NewLine) & tableAliasTemp) + Environment.NewLine + "AS " + Environment.NewLine + "(" + Environment.NewLine + Ribbon.AppVariables.ScriptRange + ") " + Environment.NewLine + "MERGE ") & tableAlias) + " AS T" + Environment.NewLine + "USING ") & tableAliasTemp) + " AS S" + Environment.NewLine + "ON " + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "WHEN NOT MATCHED BY TARGET" + Environment.NewLine + "THEN INSERT" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "VALUES" + Environment.NewLine + "(" + Environment.NewLine + Ribbon.ConcatenateColumnNames(tbl.Range, "S", "[", "]") + Environment.NewLine + ")" + Environment.NewLine + "WHEN MATCHED" + Environment.NewLine + "THEN UPDATE SET" + Environment.NewLine + Ribbon.ConcatenateColumnNamesJoin(tbl.Range, "T", "S") + "--WHEN NOT MATCHED BY SOURCE AND 'ADD WHERE CLAUSE HERE'" + Environment.NewLine + "--THEN DELETE" + Environment.NewLine + "OUTPUT $action, inserted.*, deleted.*;" + Environment.NewLine + Environment.NewLine + "ROLLBACK TRANSACTION;" + Environment.NewLine + "--COMMIT TRANSACTION;" + Environment.NewLine + "GO"
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub TSqlSelectUnion()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                sqlColName = My.Settings.Table_ColumnName

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                ' Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                        'DO NOTHING - because the column is hidden or the last column with the sql script
                        'if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        Dim valueTsuffix As String = "& "" AS [" + col.Name + "] """
                        formula += Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """") & valueTsuffix
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim firstRowNbr As Integer = tbl.Range(1, 1).Row + 1
                ' must use the offset for the first row number
                formula = (Convert.ToString("=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, """", ""UNION "") & " + """SELECT "" & ") & formula) + " & "" """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub TSqlSelectValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty

                sqlColName = (Convert.ToString("SELECT ") & tableAlias) + ".*" + " FROM (VALUES"

                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)

                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                ' Columns formatted as text will not work as formulas and the added column will copy the formatting from the previous column so ensure that the added column never has Text format...
                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                        'DO NOTHING - because the column is hidden or the last column with the sql script
                        'if (col.Name != lastColumnName | col.Range.EntireColumn.Hidden == false)
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "", "" & ")
                        End If
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"
                        formula += (Convert.ToString((Convert.ToString((Convert.ToString("""") & qt) + """ & ") & colRef) + " & """) & qt) + """"
                    End If
                Next
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                Dim firstRowNbr As Integer = tbl.Range(1, 1).Row + 1
                ' must use the offset for the first row number
                formula = (Convert.ToString("=IF(" + (firstRowNbr).ToString() + "-ROW() = 0, "" "", "","") & " + """ ( "" & ") & formula) + " & "")"""
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                tbl.ShowTotals = True
                Dim totalsColumnValue As String = (Convert.ToString(") ") & tableAlias) + " (" + Ribbon.ConcatenateColumnNames(tbl.Range, "", "[", "]") + ") "
                tbl.TotalsRowRange(lastColumnIndex).Value2 = totalsColumnValue
                ' totals row has a maximum limit of 32,767 characters
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try
        End Sub

        Public Shared Sub TSqlUpdateValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty
                Dim afterWhere As Boolean = False

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim columnName As String = col.Name
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)
                        If [String].Compare(firstPart.Trim(), "WHERE", True) = 0 Then
                            Dim firstWhereColumn As String = col.Name
                            firstWhereColumn = System.Text.RegularExpressions.Regex.Replace(firstWhereColumn, "where ", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                            firstWhereColumn = firstWhereColumn.Trim()
                            valuePrefix = (Convert.ToString("""WHERE [") & firstWhereColumn) + "] = "" & "
                            afterWhere = True
                        Else
                            If afterWhere = True Then
                                valuePrefix = """AND [" + col.Name + "] = "" & "
                            Else
                                Dim useComma As String = String.Empty
                                If columnCount <> 1 Then
                                    useComma = ","
                                End If
                                valuePrefix = (Convert.ToString("""") & useComma) + " [" + col.Name + "] = "" & "
                            End If
                        End If
                        If Ribbon.GetSqlDataType(col) = My.Settings.Column_TypeDate Then
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""DATE(")) & qt) + """ & ") & colRef) + " & """) & qt) + ", ") & qt) + My.Settings.Table_ColumnFormatDate) & qt) + ")"""
                        Else
                            formula += (Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """"
                        End If
                    End If
                Next

                Dim whereCheck As String = String.Empty
                If afterWhere = False Then
                    whereCheck = " & "" WHERE "" "
                End If
                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                formula = (Convert.ToString(Convert.ToString((Convert.ToString("=""UPDATE ") & tableAlias) + " SET"" & ") & formula) & whereCheck) + " & CHAR(10) & ""GO """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "SQL"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

        Public Shared Sub XmlValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Dim formula As String = String.Empty
            Try
                Dim lastColumnName As String = My.Settings.Table_ColumnTableAlias
                Dim tableAlias As String = My.Settings.Table_ColumnTableAlias
                Dim sqlColName As String = String.Empty
                Dim columnCount As Integer = 0

                sqlColName = My.Settings.Table_ColumnName
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                Dim lastColumnIndex As Integer = tbl.Range.Columns.Count
                sqlCol = tbl.ListColumns(lastColumnIndex)
                If sqlCol.Name = sqlColName Then
                    lastColumnName = sqlCol.Name
                Else
                    sqlCol = tbl.ListColumns.Add()
                    sqlCol.Name = lastColumnName
                    lastColumnIndex = tbl.Range.Columns.Count
                End If

                sqlCol.DataBodyRange.NumberFormat = "General"
                Dim qt As String = String.Empty

                For Each col As Excel.ListColumn In tbl.ListColumns
                    If col.Name.IndexOfAny(New Char() {"["c, "]"c, """"c}) <> -1 Then
                        MessageBox.Show("Please remove one of these incorrect characters in a column header" + Environment.NewLine + " [ " + Environment.NewLine + " ] " + Environment.NewLine + """ " + Environment.NewLine + "Column Name: " + col.Name, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    'DO NOTHING - because the column is hidden or the last column with the sql script
                    If col.Name = lastColumnName Or col.Range.EntireColumn.Hidden Then
                    Else
                        If Not String.IsNullOrEmpty(formula) Then
                            formula = formula & Convert.ToString(" & "" "" & ")
                        End If
                        If columnCount = 0 Then
                            Ribbon.AppVariables.FirstColumnName = col.Name
                        End If
                        columnCount += 1
                        qt = Ribbon.ApplyTextQuotes(col)
                        Dim colRef As String = Ribbon.GetColumnFormat(col).ToString()
                        colRef = colRef.Replace("'", "''")
                        colRef = colRef.Replace("#", "'#")
                        colRef = (Convert.ToString(Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & colRef) + ", " + """") & qt) + """, """) & qt) & qt) + """)"

                        Dim valuePrefix As String = String.Empty
                        Dim valueSuffix As String = String.Empty
                        Dim columnName As String = col.Name.ToLower()
                        Dim firstPart As String = columnName.Substring(0, columnName.LastIndexOf(" ") + 1)

                        valuePrefix = (Convert.ToString(" CHAR(10) & ""<") & columnName) + ">"" & "
                        valueSuffix = (Convert.ToString(" & ""</") & columnName) + ">"" "
                        formula += Convert.ToString((Convert.ToString((Convert.ToString((Convert.ToString(valuePrefix & Convert.ToString("""")) & qt) + """ & ") & colRef) + " & """) & qt) + """") & valueSuffix
                    End If
                Next

                Dim nullValue As String = My.Settings.Table_ColumnScriptNull
                formula = (Convert.ToString((Convert.ToString((Convert.ToString("SUBSTITUTE(") & formula) + ", ""'") & nullValue) + "'"", """) & nullValue) + """)"
                formula = (Convert.ToString("=""<row> "" & ") & formula) + " & CHAR(10) & ""</row> """
                tbl.ShowTotals = False
                lastColumnName = sqlColName
                ' maximum header characters are 255
                tbl.HeaderRowRange(lastColumnIndex).Value2 = lastColumnName
                Try
                    sqlCol.DataBodyRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Rows.Formula = formula
                    sqlCol.Range.Columns.AutoFit()
                    sqlCol.Range.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.Range.WrapText = True
                    sqlCol.DataBodyRange.Copy()
                    Ribbon.AppVariables.FileType = "XML"
                    Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

                End Try
            Catch generatedExceptionName As System.OutOfMemoryException
                MessageBox.Show("The amount of records is too big", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
                If sqlCol IsNot Nothing Then
                    'Marshal.ReleaseComObject(sqlCol)
                End If
                Ribbon.OpenScriptPane()
                Logging.InsertRecordInfo(False, formula)
            End Try

        End Sub

    End Class

End Namespace
