Imports System.Drawing

Namespace Scripts
    Namespace Syntax

        Public Class SyntaxStyle

            Public Property Bold() As Boolean
                Get
                    Return m_Bold
                End Get
                Set
                    m_Bold = Value
                End Set
            End Property
            Private m_Bold As Boolean

            Public Property Italic() As Boolean
                Get
                    Return m_Italic
                End Get
                Set
                    m_Italic = Value
                End Set
            End Property
            Private m_Italic As Boolean

            Public Property Color() As Color
                Get
                    Return m_Color
                End Get
                Set
                    m_Color = Value
                End Set
            End Property
            Private m_Color As Color

            Public Sub New(color__1 As Color, bold__2 As Boolean, italic__3 As Boolean)
                Color = color__1
                Bold = bold__2
                Italic = italic__3
            End Sub

            Public Sub New(color As Color)
                Me.New(color, False, False)
            End Sub

        End Class

    End Namespace
End Namespace