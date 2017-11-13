Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Namespace Scripts
    Namespace Syntax

        Public Class SyntaxHighlighter
            ''' <summary>
            ''' Reference to the RichTextBox instance, for which the syntax highlighting is going to occur.
            ''' </summary>
            Private ReadOnly _richTextBox As RichTextBox

            ''' <summary>
            ''' Reference to the font size of the RichTextBox instance, for which the syntax highlighting is going to occur.
            ''' </summary>
            Private ReadOnly _fontSizeFactor As Integer

            ''' <summary>
            ''' Reference to the font name of the RichTextBox instance, for which the syntax highlighting is going to occur.
            ''' </summary>
            Private ReadOnly _fontName As String

            ''' <summary>
            ''' Determines whether the program is busy creating rtf for the previous modification of the text-box. It is necessary to avoid blinks when the user is typing fast.
            ''' </summary>
            Private _isDuringHighlight As Boolean

            ''' <summary>
            ''' styleGroupPairs
            ''' </summary>
            Private _styleGroupPairs As List(Of StyleGroupPair)

            ''' <summary>
            ''' patternStyles
            ''' </summary>
            Private ReadOnly _patternStyles As New List(Of PatternStyleMap)()

            ''' <summary>
            ''' SyntaxHighlighter
            ''' </summary>
            ''' <param name="richTextBox">richTextBox</param>
            Public Sub New(richTextBox As RichTextBox)
                If richTextBox Is Nothing Then
                    Throw New ArgumentNullException("richTextBox")
                End If

                _richTextBox = richTextBox

                _fontSizeFactor = Convert.ToInt32(_richTextBox.Font.Size * 2)
                _fontName = _richTextBox.Font.Name

                DisableHighlighting = False

                AddHandler _richTextBox.TextChanged, AddressOf RichTextBox_TextChanged
            End Sub

            ''' <summary>
            ''' Gets or sets a value indicating whether highlighting should be disabled or not.
            ''' If true, the user input will remain intact. If false, the rich content will be
            ''' modified to match the syntax of the currently selected language.
            ''' </summary>
            Public Property DisableHighlighting() As Boolean
                Get
                    Return m_DisableHighlighting
                End Get
                Set
                    m_DisableHighlighting = Value
                End Set
            End Property
            Private m_DisableHighlighting As Boolean

            ''' <summary>
            ''' AddPattern
            ''' </summary>
            ''' <param name="patternDefinition">patternDefinition</param>
            ''' <param name="syntaxStyle">syntaxStyle</param>
            Public Sub AddPattern(patternDefinition As PatternDefinition, syntaxStyle As SyntaxStyle)
                AddPattern((_patternStyles.Count + 1).ToString(CultureInfo.InvariantCulture), patternDefinition, syntaxStyle)
            End Sub

            ''' <summary>
            ''' AddPattern
            ''' </summary>
            ''' <param name="name">name</param>
            ''' <param name="patternDefinition">patternDefinition</param>
            ''' <param name="syntaxStyle">syntaxStyle</param>
            Public Sub AddPattern(name As String, patternDefinition As PatternDefinition, syntaxStyle As SyntaxStyle)
                If patternDefinition Is Nothing Then
                    Throw New ArgumentNullException("patternDefinition")
                End If
                If syntaxStyle Is Nothing Then
                    Throw New ArgumentNullException("syntaxStyle")
                End If
                If [String].IsNullOrEmpty(name) Then
                    Throw New ArgumentException("name must not be null or empty", "name")
                End If

                Dim existingPatternStyle = FindPatternStyle(name)

                If existingPatternStyle IsNot Nothing Then
                    Throw New ArgumentException("A pattern style pair with the same name already exists")
                End If

                _patternStyles.Add(New PatternStyleMap(name, patternDefinition, syntaxStyle))
            End Sub

            ''' <summary>
            ''' SyntaxStyle
            ''' </summary>
            ''' <returns>returns the new style</returns>
            Protected Function GetDefaultStyle() As SyntaxStyle
                Return New SyntaxStyle(_richTextBox.ForeColor, _richTextBox.Font.Bold, _richTextBox.Font.Italic)
            End Function

            ''' <summary>
            ''' PatternStyleMap
            ''' </summary>
            ''' <param name="name">name</param>
            ''' <returns>returns the style map</returns>
            Private Function FindPatternStyle(name As String) As PatternStyleMap
                Dim patternStyle = _patternStyles.FirstOrDefault(Function(p) [String].Equals(p.Name, name, StringComparison.Ordinal))
                Return patternStyle
            End Function

            ''' <summary>
            ''' Rehighlights the text-box content.
            ''' </summary>
            Public Sub ReHighlight()
                If Not DisableHighlighting Then
                    If _isDuringHighlight Then
                        Return
                    End If

                    _richTextBox.DisableThenDoThenEnable(AddressOf HighlighTextBase)
                End If
            End Sub

            ''' <summary>
            ''' RichTextBox_TextChanged
            ''' </summary>
            ''' <param name="sender">sender</param>
            ''' <param name="e">e</param>
            Private Sub RichTextBox_TextChanged(sender As Object, e As EventArgs)
                ReHighlight()
            End Sub

            ''' <summary>
            ''' IEnumerable TODO: make abstact
            ''' </summary>
            ''' <param name="text">text</param>
            ''' <returns>parsedExpressions</returns>
            Friend Function Parse(text As String) As IEnumerable(Of Expression)
                text = text.NormalizeLineBreaks(vbLf)
                Dim parsedExpressions = New List(Of Expression)() From {
                    New Expression(text, ExpressionType.None, [String].Empty)
                }

                For Each patternStyleMap As var In _patternStyles
                    parsedExpressions = ParsePattern(patternStyleMap, parsedExpressions)
                Next

                parsedExpressions = ProcessLineBreaks(parsedExpressions)
                Return parsedExpressions
            End Function

            ''' <summary>
            ''' lineBreakRegex TODO: move to child
            ''' </summary>
            Private _lineBreakRegex As Regex

            ''' <summary>
            ''' GetLineBreakRegex TODO: move to child
            ''' </summary>
            ''' <returns>lineBreakRegex</returns>
            Private Function GetLineBreakRegex() As Regex
                If _lineBreakRegex Is Nothing Then
                    _lineBreakRegex = New Regex(Regex.Escape(vbLf), RegexOptions.Compiled)
                End If

                Return _lineBreakRegex
            End Function

            ''' <summary>
            ''' List
            ''' </summary>
            ''' <param name="expressions">expressions</param>
            ''' <returns>parsedExpressions</returns>
            Private Function ProcessLineBreaks(expressions As List(Of Expression)) As List(Of Expression)
                Dim parsedExpressions = New List(Of Expression)()

                Dim regex = GetLineBreakRegex()

                For Each inputExpression As var In expressions
                    Dim lastProcessedIndex As Integer = -1

                    For Each match As var In regex.Matches(inputExpression.Content).Cast(Of Match)().OrderBy(Function(m) m.Index)
                        If match.Success Then
                            If match.Index > lastProcessedIndex + 1 Then
                                Dim nonMatchedContent As String = inputExpression.Content.Substring(lastProcessedIndex + 1, match.Index - lastProcessedIndex - 1)
                                Dim nonMatchedExpression = New Expression(nonMatchedContent, inputExpression.Type, inputExpression.Group)
                                'lastProcessedIndex = match.Index + match.Length - 1;
                                parsedExpressions.Add(nonMatchedExpression)
                            End If

                            Dim matchedContent As String = inputExpression.Content.Substring(match.Index, match.Length)
                            Dim matchedExpression = New Expression(matchedContent, ExpressionType.Newline, "line-break")
                            parsedExpressions.Add(matchedExpression)
                            lastProcessedIndex = match.Index + match.Length - 1
                        End If
                    Next

                    If lastProcessedIndex < inputExpression.Content.Length - 1 Then
                        Dim nonMatchedContent As String = inputExpression.Content.Substring(lastProcessedIndex + 1, inputExpression.Content.Length - lastProcessedIndex - 1)
                        Dim nonMatchedExpression = New Expression(nonMatchedContent, inputExpression.Type, inputExpression.Group)
                        parsedExpressions.Add(nonMatchedExpression)
                    End If
                Next

                Return parsedExpressions
            End Function

            ''' <summary>
            ''' List TODO: move to relevant child class
            ''' </summary>
            ''' <param name="patternStyleMap">patternStyleMap</param>
            ''' <param name="expressions">expressions</param>
            ''' <returns>parsedExpressions</returns>
            Private Function ParsePattern(patternStyleMap As PatternStyleMap, expressions As List(Of Expression)) As List(Of Expression)
                Dim parsedExpressions = New List(Of Expression)()

                For Each inputExpression As var In expressions
                    If inputExpression.Type <> ExpressionType.None Then
                        parsedExpressions.Add(inputExpression)
                    Else
                        Dim regex = patternStyleMap.PatternDefinition.Regex

                        Dim lastProcessedIndex As Integer = -1

                        For Each match As var In regex.Matches(inputExpression.Content).Cast(Of Match)().OrderBy(Function(m) m.Index)
                            If match.Success Then
                                If match.Index > lastProcessedIndex + 1 Then
                                    Dim nonMatchedContent As String = inputExpression.Content.Substring(lastProcessedIndex + 1, match.Index - lastProcessedIndex - 1)
                                    Dim nonMatchedExpression = New Expression(nonMatchedContent, ExpressionType.None, [String].Empty)
                                    'lastProcessedIndex = match.Index + match.Length - 1;
                                    parsedExpressions.Add(nonMatchedExpression)
                                End If

                                Dim matchedContent As String = inputExpression.Content.Substring(match.Index, match.Length)
                                Dim matchedExpression = New Expression(matchedContent, patternStyleMap.PatternDefinition.ExpressionType, patternStyleMap.Name)
                                parsedExpressions.Add(matchedExpression)
                                lastProcessedIndex = match.Index + match.Length - 1
                            End If
                        Next

                        If lastProcessedIndex < inputExpression.Content.Length - 1 Then
                            Dim nonMatchedContent As String = inputExpression.Content.Substring(lastProcessedIndex + 1, inputExpression.Content.Length - lastProcessedIndex - 1)
                            Dim nonMatchedExpression = New Expression(nonMatchedContent, ExpressionType.None, [String].Empty)
                            parsedExpressions.Add(nonMatchedExpression)
                        End If
                    End If
                Next

                Return parsedExpressions
            End Function

            ''' <summary>
            ''' IEnumerable TODO: make abstract
            ''' </summary>
            ''' <returns>StyleGroupPair</returns>
            Friend Function GetStyles() As IEnumerable(Of StyleGroupPair)
                yield Return New StyleGroupPair(GetDefaultStyle(), [String].Empty)

		For Each patternStyle As var In _patternStyles
                    Dim style = patternStyle.SyntaxStyle
                    yield Return New StyleGroupPair(New SyntaxStyle(style.Color, style.Bold, style.Italic), patternStyle.Name)
		Next
            End Function

            ''' <summary>
            ''' GetGroupName TODO: make virtual
            ''' </summary>
            ''' <param name="expression">expression</param>
            ''' <returns>expression.Group</returns>
            Friend Overridable Function GetGroupName(expression As Expression) As String
                Return expression.Group
            End Function

            ''' <summary>
            ''' List
            ''' </summary>
            ''' <returns>styleGroupPairs</returns>
            Private Function GetStyleGroupPairs() As List(Of StyleGroupPair)
                If _styleGroupPairs Is Nothing Then
                    _styleGroupPairs = GetStyles().ToList()

                    For i As Integer = 0 To _styleGroupPairs.Count - 1
                        _styleGroupPairs(i).Index = i + 1
                    Next
                End If

                Return _styleGroupPairs
            End Function

#Region "RTF Stuff"
            ''' <summary>
            ''' The base method that highlights the text-box content.
            ''' </summary>
            Private Sub HighlighTextBase()
                _isDuringHighlight = True

                Try
                    Dim sb = New StringBuilder()

                    sb.AppendLine(RTFHeader())
                    sb.AppendLine(RTFColorTable())
                    sb.Append("\viewkind4\uc1\pard\f0\fs").Append(_fontSizeFactor).Append(" ")

                    For Each exp As var In Parse(_richTextBox.Text)
                        If exp.Type = ExpressionType.Whitespace Then
                            Dim wsContent As String = exp.Content
                            sb.Append(wsContent)
                        ElseIf exp.Type = ExpressionType.Newline Then
                            sb.AppendLine("\par")
                        Else
                            Dim content As String = exp.Content.Replace("\", "\\").Replace("{", "\{").Replace("}", "\}")

                            Dim styleGroups = GetStyleGroupPairs()

                            Dim groupName As String = GetGroupName(exp)

                            Dim styleToApply = styleGroups.FirstOrDefault(Function(s) [String].Equals(s.GroupName, groupName, StringComparison.Ordinal))

                            If styleToApply IsNot Nothing Then
                                Dim opening As String = [String].Empty, cloing As String = [String].Empty

                                If styleToApply.SyntaxStyle.Bold Then
                                    opening += "\b"
                                    cloing += "\b0"
                                End If

                                If styleToApply.SyntaxStyle.Italic Then
                                    opening += "\i"
                                    cloing += "\i0"
                                End If

                                sb.AppendFormat("\cf{0}{2} {1}\cf0{3} ", styleToApply.Index, content, opening, cloing)
                            Else
                                sb.AppendFormat("\cf{0} {1}\cf0 ", 1, content)
                            End If
                        End If
                    Next

                    sb.Append("\par }")

                    _richTextBox.Rtf = sb.ToString()
                Finally
                    _isDuringHighlight = False
                End Try
            End Sub

            ''' <summary>
            ''' RTFColorTable
            ''' </summary>
            ''' <returns>sbRtfColorTable.ToString()</returns>
            Private Function RTFColorTable() As String
                Dim styleGroupPairs = GetStyleGroupPairs()

                If styleGroupPairs.Count <= 0 Then
                    styleGroupPairs.Add(New StyleGroupPair(GetDefaultStyle(), [String].Empty))
                End If

                Dim sbRtfColorTable = New StringBuilder()
                sbRtfColorTable.Append("{\colortbl ;")

                For Each styleGroup As var In styleGroupPairs
                    sbRtfColorTable.AppendFormat("{0};", ColorUtils.ColorToRtfTableEntry(styleGroup.SyntaxStyle.Color))
                Next

                sbRtfColorTable.Append("}")

                Return sbRtfColorTable.ToString()
            End Function

            ''' <summary>
            ''' RTFHeader
            ''' </summary>
            ''' <returns>String.Concat</returns>
            Private Function RTFHeader() As String
                Return [String].Concat("{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 ", _fontName, ";}}")
            End Function

#End Region

        End Class

    End Namespace
End Namespace