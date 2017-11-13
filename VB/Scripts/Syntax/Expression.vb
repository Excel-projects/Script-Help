
Namespace Scripts
    Namespace Syntax

        Public Class Expression

            Public Property Type() As ExpressionType
                Get
                    Return m_Type
                End Get
                Private Set
                    m_Type = Value
                End Set
            End Property
            Private m_Type As ExpressionType

            Public Property Content() As String
                Get
                    Return m_Content
                End Get
                Private Set
                    m_Content = Value
                End Set
            End Property
            Private m_Content As String

            Public Property Group() As String
                Get
                    Return m_Group
                End Get
                Private Set
                    m_Group = Value
                End Set
            End Property
            Private m_Group As String

            Public Sub New(content__1 As String, type__2 As ExpressionType, group__3 As String)
                If content__1 Is Nothing Then
                    Throw New ArgumentNullException("content")
                End If
                If group__3 Is Nothing Then
                    Throw New ArgumentNullException("group")
                End If

                Type = type__2
                Content = content__1
                Group = group__3
            End Sub

            Public Sub New(content As String, type As ExpressionType)
                Me.New(content, type, [String].Empty)
            End Sub

            Public Overrides Function ToString() As String
                If Type = ExpressionType.Newline Then
                    Return [String].Format("({0})", Type)
                End If

                Return [String].Format("({0} --> {1}{2})", Content, Type, If(Group.Length > 0, Convert.ToString(" --> ") & Group, [String].Empty))
            End Function

        End Class

    End Namespace
End Namespace