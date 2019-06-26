Option Strict On
Option Explicit On

Imports System.Deployment.Application
Imports log4net
Imports log4net.Config

<Assembly: log4net.Config.XmlConfigurator(Watch:=True)>

Namespace Scripts
    Public Class Logging

        Public Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ErrorHandler))

        Public Shared Function GetPublishVersion() As String
            Try
                If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
                    Return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString()
                Else
                    Return "DEV"
                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex, True)

            Finally
                Logging.InsertRecordInfo()

            End Try

        End Function

        Public Shared Sub SetPath()
            Try
                XmlConfigurator.Configure()
                Dim h As log4net.Repository.Hierarchy.Hierarchy = CType(LogManager.GetRepository(), log4net.Repository.Hierarchy.Hierarchy)
                Dim logFileName As String = System.IO.Path.Combine(My.Settings.App_PathLog, My.Application.Info.AssemblyName & ".log")

                For Each a In h.Root.Appenders

                    If TypeOf a Is log4net.Appender.FileAppender Then

                        If a.Name.Equals("FileAppender") Then
                            Dim fa As log4net.Appender.FileAppender = CType(a, log4net.Appender.FileAppender)
                            fa.File = logFileName
                            fa.ActivateOptions()
                        End If
                    End If
                Next

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            Finally
                Logging.InsertRecordInfo()

            End Try

        End Sub

        Public Shared Sub InsertRecordInfo(Optional ByVal isStartup As Boolean = False, Optional ByVal message As String = "")
            Try
                If My.Settings.Visible_log_file_trace_on Or isStartup Then
                    Dim sf As New System.Diagnostics.StackFrame(1)
                    Dim caller As System.Reflection.MethodBase = sf.GetMethod()
                    Dim currentProcedure As String = (caller.Name).Trim()
                    Dim currentClass As String = caller.DeclaringType.FullName

                    Dim infoLine As String = String.Empty
                    infoLine += "|VERSION|" + GetPublishVersion()
                    infoLine += "|USER_NAME|" + Environment.UserName
                    infoLine += "|MACHINE_NAME|" + Environment.MachineName
                    infoLine += "|CLASS|" + currentClass
                    infoLine += "|PROCEDURE|" + currentProcedure

                    If message <> "" Then
                        infoLine += "|MESSAGE|" + message
                    End If

                    Logging.log.Info(infoLine)

                End If

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

        Public Shared Sub InsertRecordError(ByVal currentClass As String, ByVal currentProcedure As String, ByVal errorDescription As String)
            Try
                Dim errorLine As String = String.Empty
                errorLine += "|VERSION|" + GetPublishVersion()
                errorLine += "|USER_NAME|" + Environment.UserName
                errorLine += "|MACHINE_NAME|" + Environment.MachineName
                errorLine += "|CLASS|" + currentClass
                errorLine += "|PROCEDURE|" + currentProcedure
                errorLine += "|DESCRIPTION|" + errorDescription

                Logging.log.Error(errorLine)

            Catch ex As Exception
                ErrorHandler.DisplayMessage(ex)

            End Try

        End Sub

    End Class

End Namespace