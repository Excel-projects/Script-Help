
Namespace Scripts
    Namespace Syntax

        Public NotInheritable Class StringExtensions
            Private Sub New()
            End Sub

            Public Shared Function NormalizeLineBreaks(instance As String, preferredLineBreak As String) As String
                Return instance.Replace(vbCr & vbLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, preferredLineBreak)
            End Function

            Public Shared Function NormalizeLineBreaks(instance As String) As String
                Return NormalizeLineBreaks(instance, Environment.NewLine)
            End Function

        End Class

    End Namespace
End Namespace