Option Strict Off
Option Explicit On

Imports System.Windows.Forms

Namespace Scripts
    Friend Class Formula

        Friend Shared Sub DqlAppend()
        End Sub
        Friend Shared Sub DqlAppendLocked()
        End Sub
        Friend Shared Sub DqlCreate()
        End Sub
        Friend Shared Sub DqlTruncateAppend()
        End Sub
        Friend Shared Sub DqlUpdate()
        End Sub
        Friend Shared Sub DqlUpdateLocked()
        End Sub
        Friend Shared Sub GithubTable()
        End Sub
        Friend Shared Sub HtmlTable()
        End Sub
        Friend Shared Sub PlSqlCreateTable()
        End Sub
        Friend Shared Sub PlSqlInsertValues()
        End Sub
        Friend Shared Sub PlSqlMergeValues()
        End Sub
        Friend Shared Sub PlSqlSelectValues()
        End Sub
        Friend Shared Sub PlSqlSelectUnion()
        End Sub
        Friend Shared Sub PlSqlUpdateValues()
        End Sub
        Friend Shared Sub TSqlCreateTable()
        End Sub
        Friend Shared Sub TSqlInsertValues()
        End Sub
        Friend Shared Sub TSqlMergeValues()
        End Sub

        Public Shared Sub TSqlSelectValues()
            Dim tbl As Excel.ListObject = Nothing
            Dim sqlCol As Excel.ListColumn = Nothing
            Try
                ErrorHandler.CreateLogRecord()
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
                Dim formula As String = String.Empty
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
                    'Ribbon.AppVariables.FileType = "SQL"
                    'Ribbon.AppVariables.ScriptRange = DirectCast(Clipboard.GetData(DataFormats.Text), String)
                    'Ribbon.AppVariables.ScriptRange = Ribbon.AppVariables.ScriptRange.Replace("""", [String].Empty)
                Catch generatedExceptionName As System.Runtime.InteropServices.COMException
                    'Ribbon.AppVariables.ScriptRange = Convert.ToString("There was an issue creating the Excel formula." + Environment.NewLine + Environment.NewLine + "Formula: " + Environment.NewLine) & formula

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
                'Ribbon.OpenScriptPane()
            End Try
        End Sub

        Friend Shared Sub TSqlSelectValues2()
            Try
                Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
                If ErrorHandler.IsValidListObject() Then
                    Dim QueryType As String = "UPDATE" 'My.Settings.TSQL_QUERY_TYPE
                    Dim colName As String = My.Settings.Table_ColumnName
                    ' Adds a rightmost column, or updates an existing column, on a table
                    ' that contains a formula to calculate a TSQL VALUES clause.
                    ' The clause will include all columns to the left of the sql column and skip hidden columns
                    ' This is intended to allow for configuration of the contents of the VALUES clause via the UI.
                    ' Locate or create the column
                    Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    Dim sqlCol As Excel.ListColumn
                    sqlCol = Ribbon.GetItem(tbl.ListColumns, colName)
                    If sqlCol Is Nothing Then
                        sqlCol = tbl.ListColumns.Add
                        sqlCol.Name = colName
                    End If

                    ' Columns formatted as text will not work as formulas and the added column
                    ' will copy the formatting from the previous column so ensure that
                    ' the added column never has Text format...

                    sqlCol.DataBodyRange.NumberFormat = "General"

                    Dim formula As String = String.Empty
                    Dim col As Excel.ListColumn
                    Dim qt As String = String.Empty
                    Dim ColRef As String = String.Empty

                    For Each col In tbl.ListColumns
                        'If col.Name contains "#" or "'" or ... then exit the sub
                        If col.Name = colName Then 'Or col.Range.EntireColumn.Hidden Then 'TODO: fix this
                            'DO NOTHING
                        Else
                            If formula <> "" Then
                                formula = formula & " & "", "" & "
                            End If

                            qt = Ribbon.ApplyTextQuotes(col)
                            ColRef = Ribbon.GetColumnFormat(col).ToString
                            ColRef = ColRef.Replace("'", "''")
                            ColRef = ColRef.Replace("#", "'#")
                            formula = formula & """" & qt & """ & " & ColRef & " & """ & qt & """"
                        End If
                    Next

                    ' add substitute to string quotes off all nulls
                    formula = "SUBSTITUTE(" & formula & ", ""'" & My.Settings.Table_ColumnScriptNull & "'"", """ & My.Settings.Table_ColumnScriptNull & """)"

                    Select Case QueryType
                        Case Is = "INSERT"
                            ' add comma and brackets for sql VALUES statment
                            formula = "= "",( "" & " & formula & " & "")"""
                        Case Is = "UPDATE"
                            formula = "= ""UNION SELECT "" & " & formula & " & """""
                        Case Else
                            formula = "= "",( "" & " & formula & " & "")"""
                    End Select
                    sqlCol.DataBodyRange.Formula = formula
                    sqlCol.DataBodyRange.HorizontalAlignment = Excel.Constants.xlLeft
                    sqlCol.DataBodyRange.Interior.ColorIndex = Excel.Constants.xlNone
                    sqlCol.DataBodyRange.Copy()
                End If

            Catch ex As Exception
                Call ErrorHandler.DisplayMessage(ex)

            Finally
                Cursor.Current = System.Windows.Forms.Cursors.Arrow

            End Try
        End Sub
        Friend Shared Sub TSqlSelectUnion()
        End Sub
        Friend Shared Sub TSqlUpdateValues()
        End Sub
        Friend Shared Sub XmlValues()
        End Sub

    End Class

End Namespace
