Option Strict Off 'late binding on Excel object
Option Explicit On

'Imports System.Diagnostics
'Imports System.IO.Path
'Imports System.Runtime.InteropServices
'Imports Microsoft.Office.Tools.Ribbon
'Imports Excel = Microsoft.Office.Interop.Excel
'Imports System.Windows
Imports System.Windows.Forms

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon
    Implements Office.IRibbonExtensibility
    Private ribbon As Office.IRibbonUI

#Region "  Ribbon Events  "
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Loads the XML markup, either from an XML customization file or from XML markup embedded in the procedure, that customizes the Ribbon user interface.
    ''' </summary>
    ''' <param name="ribbonID">Represents the XML customization file</param>
    ''' <returns>A method that returns a bitmap image for the control id.</returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Load the ribbon
    ''' </summary>
    ''' <param name="ribbonUI">Represents the IRibbonUI instance that is provided by the Microsoft Office application to the Ribbon extensibility code.</param>
    ''' <remarks></remarks>
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Function GetButtonImage(ByVal control As Office.IRibbonControl) As System.Drawing.Bitmap
        '--------------------------------------------------------------------------------------------------------------------
        ' to assign a images to the controls on the ribbon in the xml file
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Select Case control.Id.ToString
                Case Is = "btnAddSqlColumnn"
                    Return My.Resources.Resources.QueryTypeTSql
                Case Else
                    Return Nothing
            End Select

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return Nothing

        End Try

    End Function

    ''' <summary>
    ''' To assign text to controls on the ribbon from the xml file
    ''' </summary>
    ''' <param name="control">Represents the object passed into the callback procedure of a control in a ribbon or another user interface that can be customized by using Office Fluent ribbon extensibility.</param>
    ''' <returns>A method that returns a string for a label. </returns>
    ''' <remarks></remarks>
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
                    Dim AppVersion As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
                    Return My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & AppVersion
                Case Is = "txtInstallDate"
                    Dim dteCreateDate As DateTime = System.IO.File.GetLastWriteTime(My.Application.Info.DirectoryPath.ToString & "\" & My.Application.Info.AssemblyName.ToString & ".dll") 'get creation date 
                    Return dteCreateDate.ToString("dd-MMM-yyyy hh:mm tt")
                Case Else
                    Return String.Empty
            End Select

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return String.Empty

        End Try

    End Function

    Public Sub GetSelectedItemID(ByVal Control As Office.IRibbonControl, ByRef itemID As Object)
        '--------------------------------------------------------------------------------------------------------------------
        ' to assign text to controls on the ribbon from the xml file
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Select Case Control.Id.ToString
                Case Is = "drpQueryType"
                    itemID = My.Settings.TSQL_QUERY_TYPE
                    'itemID = Range(My.Settings.TSQL_QUERY_TYPE).Value2
                Case Else
                    itemID = String.Empty
            End Select

        Catch ex As Exception
            Call ErrorMsg(ex)
            itemID = String.Empty

        End Try

    End Sub

    Public Function GetCount(ByVal Control As Office.IRibbonControl) As Integer
        '--------------------------------------------------------------------------------------------------------------------
        ' to assign how many items for the control
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Select Case Control.Id.ToString
                Case Is = "drpQueryType"
                    Return 2
                Case Else
                    Return 0
            End Select

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return 0

        End Try

    End Function

    Public Function GetLabel(ByVal Control As Office.IRibbonControl, Index As Integer) As String
        '--------------------------------------------------------------------------------------------------------------------
        ' to assign text to control items index
        '--------------------------------------------------------------------------------------------------------------------
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
            Call ErrorMsg(ex)
            Return String.Empty

        End Try

    End Function

    Public Sub MyAction(ByVal Control As Office.IRibbonControl, ItemId As String, Index As Integer)
        '--------------------------------------------------------------------------------------------------------------------
        ' to preform an action based on what item index the user selects from the control
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Select Case Control.Id.ToString
                Case Is = "drpQueryType"
                    Select Case Index
                        Case 0
                            My.Settings.TSQL_QUERY_TYPE = "INSERT"
                        Case 1
                            My.Settings.TSQL_QUERY_TYPE = "UPDATE"
                        Case Else
                            My.Settings.TSQL_QUERY_TYPE = String.Empty
                    End Select
                Case Else
                    'nothing
            End Select

        Catch ex As Exception
            Call ErrorMsg(ex)
            My.Settings.TSQL_QUERY_TYPE = String.Empty

        End Try

    End Sub

    Public Function GetVisible(ByVal Control As Office.IRibbonControl) As Boolean
        '--------------------------------------------------------------------------------------------------------------------
        ' to assign the visiblity to controls
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Select Case Control.Id.ToString
                Case Is = "btnForceColumnToDate"
                    Return My.Settings.Visible_btnForceColumnToDate
                Case Is = "drpQueryType"
                    Return My.Settings.Visible_drpQueryType
                Case Is = "ComAddInsDialog"
                    Return My.Settings.Visible_ComAddInsDialog
                Case Is = "FormatAsTableGallery"
                    Return My.Settings.Visible_FormatAsTableGallery
                Case Is = "ViewFreezePanesGallery"
                    Return My.Settings.Visible_ViewFreezePanesGallery
                Case Is = "RemoveDuplicates"
                    Return My.Settings.Visible_RemoveDuplicates
                Case Else
                    Return False
            End Select

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return False

        End Try

    End Function

    Public Function GetEnabled(ByVal control As Office.IRibbonControl) As Boolean
        Try
            Return True

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return False

        End Try

    End Function

#End Region

#Region "  Ribbon Buttons  "

    Public Sub ForceColumnToDate(ByVal control As Office.IRibbonControl)
        Dim col As Excel.ListColumn = ListColumn(Globals.ThisAddIn.Application.ActiveCell)
        Try
            ' Helper to fix the problem of the default date (time) format selected by
            ' Excel when pasting data from SSMS. There does not seem to be any way to change this.
            ' See: http://superuser.com/questions/552285/how-do-i-change-the-default-format-for-an-date-time-value-copied-from-ssms-to-ex

            If (col Is Nothing) Then
                Exit Try
            End If

            col.DataBodyRange.Cells(1).Activate()

            Dim cell As Excel.Range
            For Each cell In col.DataBodyRange
                If IsDate(cell.Value) Then
                    cell.Value = CDate(cell.Value)
                End If
            Next

        Catch ex As Exception
            Call ErrorMsg(ex)

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(col)

        End Try

    End Sub

    Public Sub FormatSqlDateColumns(ByVal control As Office.IRibbonControl)
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose: Finds dates columns with SSMS crap format and alters to use standard TSQL date format
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
            If IsValidListObject(tbl) Then
                Dim col As Excel.ListColumn
                Dim cell As Excel.Range
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                For Each col In tbl.ListColumns
                    cell = FirstNotNull(col.DataBodyRange)
                    If (Not cell Is Nothing) Then
                        If cell.NumberFormat.ToString = My.Settings.TSQL_DATE_PASTE_FORMAT Or IsDate(cell.Value) Then
                            col.DataBodyRange.NumberFormat = My.Settings.TSQL_DATE_FORMAT
                            col.DataBodyRange.HorizontalAlignment = Excel.Constants.xlCenter
                        End If
                    End If
                Next
                tbl.DataBodyRange.Interior.ColorIndex = Excel.Constants.xlNone
            End If

        Catch ex As Exception
            Call ErrorMsg(ex)

        Finally
            Cursor.Current = System.Windows.Forms.Cursors.Arrow

        End Try

    End Sub

    Public Sub CleanData(ByVal control As Office.IRibbonControl)
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose: Clean all text data and returns number of cells altered
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
            If IsValidListObject(tbl) Then
                Dim sqlColName As String = My.Settings.SQL_COL_NAME
                Dim col As Excel.ListColumn
                Dim cell As Excel.Range
                Dim a As Object '(,) As String
                Dim c As String = String.Empty
                Dim cc As String = String.Empty
                Dim cnt As Integer = 0
                Dim i As Integer = 0
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                For Each col In tbl.ListColumns
                    a = col.DataBodyRange.Value2
                    'For i = LBound(a) To UBound(a)
                    For i = LBound(CType(a, Array)) To UBound(CType(a, Array)) 'TODO: fix error, if table only has 1 row, for now added InvalidCastException  -- For i = LBound(a) To UBound(a) --For i = 0 To a.GetUpperBound(0) 
                        c = a(i, 1)
                        If Globals.ThisAddIn.Application.WorksheetFunction.IsText(c) Then
                            cc = Trim(Globals.ThisAddIn.Application.WorksheetFunction.Clean(c.ToString))
                            If (cc <> c) Then
                                cell = col.DataBodyRange.Cells(1).Offset(i - 1, 0)
                                cell.Value = cc
                                cell.Interior.Color = My.Settings.CLEAN_CELL_COLOUR
                                cnt = cnt + 1
                            End If
                        End If
                    Next
                Next
                'Globals.ThisAddIn.Application.StatusBar = "Cleaning data..."
                'Globals.ThisAddIn.Application.StatusBar = "Number of cells cleaned: " & cnt.ToString
                MessageBox.Show("The number of cells cleaned: " & cnt.ToString, "Cleaning has finished", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As System.InvalidCastException 'TODO: fix error, if table only has 1 row
            MessageBox.Show("Please insert one more row.", "Unable to clean 1 row", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            Call ErrorMsg(ex)

        Finally
            Cursor.Current = System.Windows.Forms.Cursors.Arrow

        End Try

    End Sub

    Public Sub AddSqlColumnn(ByVal control As Office.IRibbonControl)
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose: Add a formula at the end of the table to use in an insert statement
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Dim tbl As Excel.ListObject = Globals.ThisAddIn.Application.ActiveCell.ListObject
            If IsValidListObject(tbl) Then
                Dim QueryType As String = My.Settings.TSQL_QUERY_TYPE
                Dim colName As String = My.Settings.SQL_COL_NAME
                ' Adds a rightmost column, or updates an existing column, on a table
                ' that contains a formula to calculate a TSQL VALUES clause.
                ' The clause will include all columns to the left of the sql column and skip hidden columns
                ' This is intended to allow for configuration of the contents of the VALUES clause via the UI.
                ' Locate or create the column
                Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                Dim sqlCol As Excel.ListColumn
                sqlCol = GetItem(tbl.ListColumns, colName)
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
                    If col.Name = colName Or col.Range.EntireColumn.Hidden Then
                        'DO NOTHING
                    Else
                        If formula <> "" Then
                            formula = formula & " & "", "" & "
                        End If

                        qt = ColumnQuote(col)
                        ColRef = ColumnReference(col).ToString
                        ColRef = ColRef.Replace("'", "''")
                        ColRef = ColRef.Replace("#", "'#")
                        formula = formula & """" & qt & """ & " & ColRef & " & """ & qt & """"
                    End If
                Next

                ' add substitute to string quotes off all nulls
                formula = "SUBSTITUTE(" & formula & ", ""'" & My.Settings.TSQL_NULL & "'"", """ & My.Settings.TSQL_NULL & """)"

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
            Call ErrorMsg(ex)

        Finally
            Cursor.Current = System.Windows.Forms.Cursors.Arrow

        End Try

    End Sub

    Public Sub OpenSettingsForm(ByVal control As Office.IRibbonControl)
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose: show the settings form
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Dim FormSettings As New frmSettings
            FormSettings.ShowDialog()
            ribbon.Invalidate()

        Catch ex As Exception
            Call ErrorMsg(ex)

        End Try

    End Sub

    Public Sub OpenHelpFile(ByVal control As Office.IRibbonControl)
        Call OpenFile(My.Settings.HelpFile)
    End Sub

#End Region

#Region "  Subroutines  "

    Private Function ColumnQuote(ByVal col As Excel.ListColumn) As String
        Try
            ' Some columns in TSQL will need quoting and others will not
            If (SqlType(col) <> My.Settings.TSQL_NUMERIC) Then
                Return My.Settings.TSQL_QUOTE
            Else
                Return String.Empty
            End If

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return String.Empty

        End Try

    End Function

    Private Function ColumnReference(ByVal col As Excel.ListColumn) As String
        Try
            Dim fmt As String = String.Empty
            Dim nFmt As String = String.Empty

            Select Case SqlType(col)
                Case My.Settings.TSQL_DATE
                    fmt = My.Settings.TSQL_DATE_FORMAT

                Case My.Settings.TSQL_NUMERIC
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
            Call ErrorMsg(ex)
            Return String.Empty

        End Try

    End Function

    Private Function Formatted(ByVal col As Excel.ListColumn, fmt As String) As String
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose: Generate a formula reference with optional text formatting
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Formatted = "[" & col.Name & "]"
            If (fmt = "") Then
                Exit Try
            End If
            Return "TEXT(" & Formatted & ",""" & fmt & """)"

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return String.Empty

        End Try

    End Function

    Private Function SqlType(ByVal col As Excel.ListColumn) As Integer
        Try
            ' Determine the likely SQL type of the column
            ' default to text
            SqlType = My.Settings.TSQL_TEXT
            Dim rowCnt As Integer = col.DataBodyRange.Rows.Count
            Dim numCnt As Double = 0
            Dim notNullCnt As Double = Globals.ThisAddIn.Application.WorksheetFunction.CountIf(col.DataBodyRange, "<>" & My.Settings.TSQL_NULL)

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
            If IsDate(FirstNotNull(col.DataBodyRange)) Or col.DataBodyRange.NumberFormat.ToString = My.Settings.TSQL_DATE_FORMAT Then
                Return My.Settings.TSQL_DATE
            Else
                Return My.Settings.TSQL_NUMERIC
            End If

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return 0

        End Try

    End Function

    Public Function GetItem(ByVal col As Excel.ListColumns, key As String) As Excel.ListColumn
        Try
            Return col(key)

        Catch ex As Exception
            'Call ErrorMsg(ex)            
            Return Nothing

        End Try

    End Function

    Private Function ListColumn(ByVal cell As Excel.Range) As Excel.ListColumn
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
            Call ErrorMsg(ex)
            Return Nothing
            Exit Try

        End Try

    End Function

    Private Function FirstNotNull(ByVal rng As Excel.Range) As Excel.Range
        ' TODO: find a way to do this without looping.
        ' NOTE: SpecialCells is unreliable when called from VBA UDFs (Odd ??!)        
        Try
            If (rng Is Nothing) Then
                Return Nothing
            End If
            Dim cell As Excel.Range

            For Each cell In rng
                If (cell.Value IsNot Nothing) Then
                    If (cell.Value.ToString <> My.Settings.TSQL_NULL) Then
                        Return cell
                    End If
                End If
            Next
            Return Nothing

        Catch ex As Exception
            Call ErrorMsg(ex)
            Return Nothing

        End Try

    End Function

    Public Function IsValidListObject(ByVal tbl As Excel.ListObject) As Boolean
        Try
            If (tbl Is Nothing) Then
                MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again.", My.Application.Info.Description.ToString, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            'Call ErrorMsg(ex)
            Return False

        End Try

    End Function

    Public Sub ErrorMsg(ByRef ex As Exception)
        '--------------------------------------------------------------------------------------------------------------------
        ' Global error message for all procedures
        '--------------------------------------------------------------------------------------------------------------------
        Dim Msg As String
        Dim sf As New System.Diagnostics.StackFrame(1)
        Dim caller As System.Reflection.MethodBase = sf.GetMethod()
        Dim Procedure As String = (caller.Name).Trim

        Msg = "Contact your system administrator." & vbCrLf
        Msg += "Procedure: " & Procedure & vbCrLf
        Msg += "Description: " & ex.ToString & vbCrLf   '
        MsgBox(Msg, vbCritical, "Unexpected Error")

    End Sub

    Public Sub OpenFile(ByVal FilePath As String)
        '--------------------------------------------------------------------------------------------------------------------
        ' Purpose: open a file from the source list
        '--------------------------------------------------------------------------------------------------------------------
        Try
            Dim pStart As New System.Diagnostics.Process
            If FilePath = String.Empty Then Exit Try
            pStart.StartInfo.FileName = FilePath
            pStart.Start()

        Catch ex As System.ComponentModel.Win32Exception
            'MessageBox.Show("No application is assicated to this file type." & vbCrLf & vbCrLf & FilePath, "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Try

        Catch ex As Exception
            Call ErrorMsg(ex)

            'Finally
            '    pStart = Nothing

        End Try

    End Sub

#End Region

End Class