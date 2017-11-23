Option Strict On
Option Explicit On

Imports System.Environment
Imports System.Windows.Forms

Namespace Scripts

    Public Class ErrorHandler

        Public Shared Sub CreateLogRecord()
            Try
                Dim sf As New System.Diagnostics.StackFrame(1)
                Dim caller As System.Reflection.MethodBase = sf.GetMethod()
                Dim currentProcedure As String = (caller.Name).Trim()
                'log.Info((Convert.ToString("[PROCEDURE]=|") & currentProcedure) + "|[USER NAME]=|" + Environment.UserName + "|[MACHINE NAME]=|" + Environment.MachineName)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
            End Try

        End Sub

        Public Shared Sub DisplayMessage(ex As Exception, Optional isSilent As [Boolean] = False)
            Dim sf As New System.Diagnostics.StackFrame(1)
            Dim caller As System.Reflection.MethodBase = sf.GetMethod()
            Dim currentProcedure As String = (caller.Name).Trim()
            Dim currentFileName As String = "" 'AssemblyInfo.GetCurrentFileName()
            Dim errorMessageDescription As String = ex.ToString()
            errorMessageDescription = System.Text.RegularExpressions.Regex.Replace(errorMessageDescription, "\r\n+", " ")
            Dim msg As String = "Contact your system administrator. A record has been created in the log file." + Environment.NewLine
            msg += (Convert.ToString("Procedure: ") & currentProcedure) + Environment.NewLine
            msg += "Description: " + ex.ToString() + Environment.NewLine
            'log.Error("[PROCEDURE]=|" + currentProcedure + "|[USER NAME]=|" + Environment.UserName + "|[MACHINE NAME]=|" + Environment.MachineName + "|[DESCRIPTION]=|" + errorMessageDescription)
            If isSilent = False Then
                MessageBox.Show(msg, "Unexpected Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
            End If

        End Sub

        Public Shared Function IsActiveDocument(Optional showMsg As Boolean = False) As Boolean
            Try
                If Globals.ThisAddIn.Application.ActiveWorkbook Is Nothing Then
                    If showMsg = True Then
                        MessageBox.Show("The command could not be completed.  Please open a document and select a range.", My.Application.Info.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    Return False
                Else
                    Return True
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False
            End Try

        End Function

        Public Shared Function IsActiveSelection(Optional showMsg As Boolean = False) As Boolean
            Dim checkRange As Excel.Range = Nothing
            Try
                checkRange = TryCast(Globals.ThisAddIn.Application.Selection, Excel.Range)
                'must cast the selection as range or errors
                If checkRange Is Nothing Then
                    If showMsg = True Then
                        MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again. [Range]", My.Application.Info.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    Return False
                Else
                    Return True
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False

            Finally
                If checkRange IsNot Nothing Then
                    'Marshal.ReleaseComObject(checkRange)
                End If
            End Try

        End Function

        Public Shared Function IsValidListObject(Optional showMsg As Boolean = False) As Boolean
            Dim tbl As Excel.ListObject = Nothing
            Try
                tbl = Globals.ThisAddIn.Application.ActiveCell.ListObject
                ' directly after the table is created this is not true
                If (tbl Is Nothing) Then
                    If showMsg = True Then
                        MessageBox.Show("The command could not be completed by using the range specified.  Select a single cell within the range and try the command again. [ListObject]", My.Application.Info.Description, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    Return False
                Else
                    Return True
                End If

            Catch generatedExceptionName As Exception
                Return False

            Finally
                If tbl IsNot Nothing Then
                    'Marshal.ReleaseComObject(tbl)
                End If
            End Try

        End Function

        Private Shared Function IsInCellEditingMode(Optional showMsg As Boolean = False) As Boolean
            Dim flag As Boolean = False
            Try
                'This will throw an Exception if Excel is in Cell Editing Mode
                Globals.ThisAddIn.Application.DisplayAlerts = False

            Catch generatedExceptionName As Exception
                If showMsg = True Then
                    MessageBox.Show("The procedure can not run while you are editing a cell.", "No action taken.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
                flag = True
            End Try
            Return flag

        End Function

        Public Shared Function IsEnabled(Optional showMsg As Boolean = False) As Boolean
            Try
                If IsActiveDocument(showMsg) = False Then
                    Return False
                Else
                    If IsActiveSelection(showMsg) = False Then
                        Return False
                    Else
                        If IsInCellEditingMode(showMsg) = True Then
                            Return False
                        Else
                            Return True
                        End If
                    End If
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False
            End Try

        End Function

        Public Shared Function IsAvailable(Optional showMsg As Boolean = False) As Boolean
            Try
                If IsEnabled(showMsg) = False Then
                    Return False
                Else
                    If IsValidListObject(showMsg) = False Then
                        Return False
                    Else
                        Return True
                    End If
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)
                Return False
            End Try

        End Function

        Public Shared Function IsDate(expression As Object) As Boolean
            If expression IsNot Nothing Then
                If TypeOf expression Is DateTime Then
                    Return True
                End If

                If TypeOf expression Is String Then
                    Dim time1 As DateTime
                    Return DateTime.TryParse(DirectCast(expression, String), time1)
                End If

            End If

            Return False

        End Function

    End Class

End Namespace
