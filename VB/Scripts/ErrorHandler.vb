Option Strict On
Option Explicit On

Imports System.Environment
Imports System.Windows.Forms

Namespace Scripts

	Module ErrorHandler

        ''' <summary> 
        ''' Global error message for all procedures
        ''' </summary>
        ''' <param name="ex">the handled exception</param>
        Public Sub DisplayMessage(ByRef ex As Exception)
            Dim caption As String = "Unexpected Error"
            Dim sf As New System.Diagnostics.StackFrame(1)
            Dim caller As System.Reflection.MethodBase = sf.GetMethod()
            Dim procedure As String = (caller.Name).Trim
            Dim msg As String = "Contact your system administrator."
            msg += NewLine & "Procedure: " & procedure
            msg += NewLine & "Description: " & ex.ToString
            'Console.WriteLine(msg)
            MessageBox.Show(msg, caption, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Sub

    End Module

End Namespace
