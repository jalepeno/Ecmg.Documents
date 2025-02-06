'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ApplicationLogging_Serilog.vb
'   Description :  [type_description_here]
'   Created     :  2/6/2025 10:43:53 AM
'   <copyright company="Conteage Corp.">
'       Copyright (c) Conteage Corp.. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports Serilog
Imports Serilog.Events
Imports Serilog.Formatting.Json

#End Region

Namespace Utilities

  Partial Public Class ApplicationLogging

    Private Shared Sub WriteSerilogEntry(ByVal lpMessage As String, lpEventType As TraceEventType)
      Dim lenuLogEventLevel As LogEventLevel = TranslateTraceLevel(lpEventType)

      Select Case lenuLogEventLevel
        Case LogEventLevel.Fatal
          Log.Fatal(lpMessage)
        Case LogEventLevel.Error
          Log.Error(lpMessage)
        Case LogEventLevel.Warning
          Log.Warning(lpMessage)
        Case LogEventLevel.Information
          Log.Information(lpMessage)
        Case LogEventLevel.Verbose
          Log.Verbose(lpMessage)
        Case Else
          Log.Information(lpMessage)
      End Select

    End Sub

    Private Shared Sub InitializeSerilog(ByVal lpApplicationName As String, ByVal lpMaxLogSize As Integer, Optional lpAsJson As Boolean = False)
      Dim lstrBaseFileName As String = GetCurrentBaseLogFile(lpApplicationName, lpAsJson)

      Log.Logger = New LoggerConfiguration().WriteTo.File(New JsonFormatter(), lstrBaseFileName, LogEventLevel.Debug, lpMaxLogSize,,,, New TimeSpan(0, 0, 30), RollingInterval.Day, lpMaxLogSize).CreateLogger()

      Dim lstrMessage As String = String.Format("{0} '{1}' Loading under user '{2}', the current locale is '{3}'",
                                          lpApplicationName,
                                          Assembly.GetExecutingAssembly.GetName().Version.ToString(),
                                          Environment.UserName,
                                          Globalization.CultureInfo.CurrentCulture.Name)

      Log.Information(lstrMessage)
    End Sub

  End Class

End Namespace