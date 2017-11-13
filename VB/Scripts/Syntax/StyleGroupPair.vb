
Namespace Scripts
    Namespace Syntax

        Friend Class StyleGroupPair

            Public Property Index() As Integer
                Get
                    Return m_Index
                End Get
                Set
                    m_Index = Value
                End Set
            End Property
            Private m_Index As Integer

            Public Property SyntaxStyle() As SyntaxStyle
                Get
                    Return m_SyntaxStyle
                End Get
                Set
                    m_SyntaxStyle = Value
                End Set
            End Property
            Private m_SyntaxStyle As SyntaxStyle

            Public Property GroupName() As String
                Get
                    Return m_GroupName
                End Get
                Set
                    m_GroupName = Value
                End Set
            End Property
            Private m_GroupName As String

            Public Sub New(syntaxStyle__1 As SyntaxStyle, groupName__2 As String)
                If syntaxStyle__1 Is Nothing Then
                    Throw New ArgumentNullException("syntaxStyle")
                End If
                If groupName__2 Is Nothing Then
                    Throw New ArgumentNullException("groupName")
                End If

                SyntaxStyle = syntaxStyle__1
                GroupName = groupName__2
            End Sub

        End Class


    End Namespace
End Namespace