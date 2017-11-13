Imports System.Collections.Generic

Namespace Scripts
    Namespace Syntax

        Public Class CaseInsensitivePatternDefinition

            Inherits PatternDefinition

            Public Sub New(tokens As IEnumerable(Of String))
                'MyBase.New(False, tokens)
            End Sub

            Public Sub New(ParamArray tokens As String())
                'MyBase.New(False, tokens)
            End Sub

        End Class

    End Namespace
End Namespace