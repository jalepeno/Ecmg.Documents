'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ApplicationLogging_Log4Net.vb
'   Description :  [type_description_here]
'   Created     :  2/6/2025 10:43:53 AM
'   <copyright company="Conteage Corp.">
'       Copyright (c) Conteage Corp.. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Reflection
Imports log4net
Imports log4net.Config
Imports Serilog.Events

#End Region

Namespace Utilities

  Partial Public Class ApplicationLogging

#Region "Class Variables"

    Private Shared _netlog As ILog

#End Region

#Region "Public Properties"

    Public Shared Property NetLog As ILog
      Get
        If _netlog Is Nothing Then
          ConfigureLog4Net()
        End If
        Return _netlog
      End Get
      Set(value As ILog)
        _netlog = value
      End Set
    End Property

#End Region

#Region "Private Methods"

    Private Shared Sub WriteLog4NetEntry(ByVal lpMessage As String, lpEventType As TraceEventType)
      Dim lenuLogEventLevel As LogEventLevel = TranslateTraceLevel(lpEventType)

      Select Case lenuLogEventLevel
        Case LogEventLevel.Fatal
          NetLog.Fatal(lpMessage)
        Case LogEventLevel.Error
          NetLog.Error(lpMessage)
        Case LogEventLevel.Warning
          NetLog.Warn(lpMessage)
        Case LogEventLevel.Information
          NetLog.Info(lpMessage)
        Case LogEventLevel.Verbose
          ' Don't log this level, there is no verbose setting for log4net
        Case Else
          NetLog.Info(lpMessage)
      End Select

    End Sub

    Private Shared Sub ConfigureLog4Net()
      Try
        Dim lobjEntryAssembly As Assembly = Assembly.GetEntryAssembly()
        _netlog = LogManager.GetLogger(lobjEntryAssembly.FullName)
        Dim mobjLogRepository As Object = LogManager.GetRepository(lobjEntryAssembly)
        XmlConfigurator.Configure(mobjLogRepository, New FileInfo("log4net.config"))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
