Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms
Imports ScriptHelp.Scripts

Public Class frmSettings

    Private Sub frmSettings_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            My.Settings.Save()

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)
            Exit Try

        End Try

    End Sub

    Private Sub frmSettings_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            Call LoadSettings()
            Call SetFormIcon(Me, My.Resources.Save)
            Dim version As String = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build & "." & My.Application.Info.Version.Revision
            Me.Text = "Settings for " & My.Application.Info.Title.ToString.Replace("&", "&&") & Space(1) & version

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)
            Exit Try

        End Try

    End Sub

    Public Sub LoadSettings()
        Try
            Me.pgdSettings.SelectedObject = My.Settings
            ''Only show "user" settings
            'Dim userAttr As New System.Configuration.UserScopedSettingAttribute
            'Dim attrs As New System.ComponentModel.AttributeCollection(userAttr)
            'pgdSettings.BrowsableAttributes = attrs

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)
            Exit Try

        End Try

    End Sub

    Public Sub SetFormIcon(ByRef frmCurrent As System.Windows.Forms.Form, ByRef bmp As Bitmap)
        Try
            frmCurrent.Icon = Icon.FromHandle(bmp.GetHicon)

        Catch ex As Exception
            ErrorHandler.DisplayMessage(ex)
            Exit Try

        End Try

    End Sub

End Class