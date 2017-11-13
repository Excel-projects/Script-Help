Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Namespace Scripts
    Namespace Syntax

        Public Class PatternDefinition
            Private ReadOnly _regex As Regex
            Private _expressionType As ExpressionType = ExpressionType.Identifier
            Private ReadOnly _isCaseSensitive As Boolean = False

            Public Sub New(regularExpression As Regex)
                If regularExpression Is Nothing Then
                    Throw New ArgumentNullException("regularExpression")
                End If
                _regex = regularExpression
            End Sub

            Public Sub New(regexPattern As String)
                If [String].IsNullOrEmpty(regexPattern) Then
                    Throw New ArgumentException("regex pattern must not be null or empty", "regexPattern")
                End If

                _regex = New Regex(regexPattern, RegexOptions.Compiled)
            End Sub

            Public Sub New(ParamArray tokens As String())
                Me.New(True, tokens)
            End Sub

            Public Sub New(tokens As IEnumerable(Of String))
                Me.New(True, tokens)
            End Sub

            Friend Sub New(caseSensitive As Boolean, tokens As IEnumerable(Of String))
                If tokens Is Nothing Then
                    Throw New ArgumentNullException("tokens")
                End If

                caseSensitive = _isCaseSensitive

                Dim regexTokens = New List(Of String)()

                For Each token As var In tokens
                    Dim escaptedToken = Regex.Escape(token.Trim())

                    If escaptedToken.Length > 0 Then
                        If [Char].IsLetterOrDigit(escaptedToken(0)) Then
                            regexTokens.Add([String].Format("\b{0}\b", escaptedToken))
                        Else
                            regexTokens.Add(escaptedToken)
                        End If
                    End If
                Next

                Dim pattern As String = [String].Join("|", regexTokens)
                Dim regexOptions__1 = RegexOptions.Compiled
                If Not caseSensitive Then
                    regexOptions__1 = regexOptions__1 Or RegexOptions.IgnoreCase
                End If
                _regex = New Regex(pattern, regexOptions__1)
            End Sub

            Friend Property ExpressionType() As ExpressionType
                Get
                    Return _expressionType
                End Get
                Set
                    _expressionType = Value
                End Set
            End Property

            Friend ReadOnly Property IsCaseSensitive() As Boolean
                Get
                    Return _isCaseSensitive
                End Get
            End Property

            Friend ReadOnly Property Regex() As Regex
                Get
                    Return _regex
                End Get
            End Property

        End Class

    End Namespace
End Namespace