Imports System.Drawing

Namespace Scripts
    Namespace Syntax

        Public Class ColorUtils

            Public Shared Function ColorToRtfTableEntry(color As Color) As String
                Return [String].Format("\red{0}\green{1}\blue{2}", color.R, color.G, color.B)
            End Function

        End Class

    End Namespace
End Namespace