
Namespace Scripts
    Namespace Syntax

        Friend Class PatternStyleMap

            Public Property Name() As String
                Get
                    Return m_Name
                End Get
                Set
                    m_Name = Value
                End Set
            End Property
            Private m_Name As String

            Public Property PatternDefinition() As PatternDefinition
                Get
                    Return m_PatternDefinition
                End Get
                Set
                    m_PatternDefinition = Value
                End Set
            End Property
            Private m_PatternDefinition As PatternDefinition

            Public Property SyntaxStyle() As SyntaxStyle
                Get
                    Return m_SyntaxStyle
                End Get
                Set
                    m_SyntaxStyle = Value
                End Set
            End Property
            Private m_SyntaxStyle As SyntaxStyle

            Public Sub New(name__1 As String, patternDefinition__2 As PatternDefinition, syntaxStyle__3 As SyntaxStyle)
                If patternDefinition__2 Is Nothing Then
                    Throw New ArgumentNullException("patternDefinition")
                End If
                If syntaxStyle__3 Is Nothing Then
                    Throw New ArgumentNullException("syntaxStyle")
                End If
                If [String].IsNullOrEmpty(name__1) Then
                    Throw New ArgumentException("name must not be null or empty", "name")
                End If

                Name = name__1
                PatternDefinition = patternDefinition__2
                SyntaxStyle = syntaxStyle__3
            End Sub

        End Class

    End Namespace
End Namespace