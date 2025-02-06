'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

'Imports MigraDoc.DocumentObjectModel
'Imports Ecmg.Cts.Configuration

#Region "Imports"

Imports System.Configuration
Imports System.Configuration.ConfigurationManager
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Documents.Configuration
Imports Serilog
Imports Serilog.Events
Imports Serilog.Formatting.Json

#End Region

Namespace Utilities

  ''' <summary>
  ''' Contains shared methods for wrting information to a log file
  ''' </summary>
  ''' <remarks></remarks>
  Partial Public Class ApplicationLogging

#Region "Class Constants"

    Private Const CTS_HKEY_PATH As String = "Software\Ecmg\Content Transformation Services"
    Private Const REG_STR_VAL_INSTALLED_PATH As String = "Installed Path"
    Private Const REG_STR_VAL_DOCS_ROOT As String = "Documents Root Path"

    Public Const OPEN_LOG_EVENT_ID As Integer = 67360
    Public Const CLOSE_LOG_EVENT_ID As Integer = 25673
    Private Const MAX_LOG_SIZE As Integer = 10485760 '10000
    Friend Const SETTING_MAX_LOG_SIZE As String = "MaxLogSize"

#End Region

#Region "Enumerations"

    Public Enum LoggingFramework
      Log4Net = 0
      Serilog = 1
    End Enum

#End Region

#Region "Class Variables"

    Private ReadOnly mstrLogDirectory As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the name of the log file for the current application
    ''' </summary>
    ''' <value></value>
    ''' <returns>The fully qualified log file name</returns>
    ''' <remarks></remarks>
    ''Public Shared ReadOnly Property LogFile() As String
    ''  Get
    ''    Try
    ''      Return GetLogFileName()
    ''    Catch ex As Exception
    ''      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    ''      'Re-throw the exception to the caller
    ''      Throw
    ''    End Try
    ''  End Get
    ''End Property

    ''' <summary>
    ''' Gets the default log file directory for application logging.
    ''' </summary>
    ''' <value></value>
    ''' <returns>
    ''' If the operating system is XP or older returns &lt;applicationpath&gt;\Logs
    ''' If the operating system is Vista or newer returns 'CurrentUserApplicationData'\Logs for non 
    ''' Cts apps and 'CurrentUserApplicationData'\Cts\Logs for Cts apps.
    ''' </returns>
    ''' <remarks>
    ''' If the registry key 
    ''' 'HKEY_LOCAL_MACHINE\Software\Ecmg\Content 
    ''' Transformation Services\Documents Root Path' 
    ''' is present then that value will be used.  
    ''' If not, then the default value will depend 
    ''' on the version of Windows.  Any version before 
    ''' 6.0 (XP, Server 2003 or older) will default to
    ''' the installed path.  This is usually something like 
    ''' 'C:\Program Files\ECMG\Content Transformation Services'.
    ''' Versions of Windows equal to or greater than 6.0 will
    ''' default to either 'My Documents\CTS' or the current 
    ''' application data path depending on whether or not the
    ''' current user is a named user or 'LOCAL SYSTEM'.
    ''' </remarks>
    Public Shared ReadOnly Property LogDirectory() As String
      Get
        Try

          Dim lstrLogDirectory As String = String.Empty
          'Dim lobjKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(CTS_HKEY_PATH)

          'If Not lobjKey Is Nothing Then
          '  ' Try to get the value of the Documents Root Path string value
          '  ' If the entry is not present in the registry this will return Nothing
          '  Dim lstrCtsDocumentsRootPath As String = lobjKey.GetValue(REG_STR_VAL_DOCS_ROOT)
          '  If lstrCtsDocumentsRootPath IsNot Nothing Then
          '    ' We found the entry, make sure it exists.
          '    If IO.Directory.Exists(lstrCtsDocumentsRootPath) = False Then
          '      Throw New IO.DirectoryNotFoundException( _
          '        String.Format("The path '{0}' specified in 'HKEY_LOCAL_MACHINE\Software\Ecmg\Content Transformation Services\Documents Root Path' is invalid.  Make sure a valid path is configured", _
          '                      lstrCtsDocumentsRootPath))
          '    End If
          '    ' Make sure it is clean.
          '    If (lstrCtsDocumentsRootPath.EndsWith("\")) = False Then
          '      lstrCtsDocumentsRootPath += "\"
          '    End If

          '    lstrLogDirectory = Helper.RemoveExtraBackSlashesFromFilePath( _
          '      String.Format("{0}Logs\", lstrCtsDocumentsRootPath))

          '  End If
          'End If

          ''Dim lstrDocumentsRootPath As String = InstallHelper.GetDocumentsRootPath
          ''lstrLogDirectory = String.Format("{0}\Logs\", lstrDocumentsRootPath)

          ''If lstrLogDirectory.Length = 0 Then

          ''  If My.Computer.Info.OSVersion > "6.0" Then
          ''    ' This is Vista or newer, access to the installed folder is limited

          ''    lstrLogDirectory = String.Format("{0}Logs\{1}",
          ''                                     Helper.GetCompanyAppDataPath,
          ''                                     My.Application.Info.AssemblyName)

          ''    If System.IO.Directory.Exists(lstrLogDirectory) = False Then
          ''      System.IO.Directory.CreateDirectory(lstrLogDirectory)
          ''    End If
          ''  Else
          ''    ' This is XP or older, we can use the installed folder
          ''    lstrLogDirectory = String.Format("{0}\Logs", My.Application.Info.DirectoryPath)
          ''  End If

          ''End If

          ''Return Helper.CleanPath(lstrLogDirectory)

          Dim lstrApplicationDirectory As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly.Location)

          lstrLogDirectory = Path.Combine(lstrApplicationDirectory, "logs")

          If Not Directory.Exists(lstrLogDirectory) Then
            Directory.CreateDirectory(lstrLogDirectory)
          End If

          Return lstrLogDirectory

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Shared ReadOnly Property BaseLogFileName As String
      Get
        Try
          Return $"{Assembly.GetExecutingAssembly.GetName.Name}.log"
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    'Public Shared Property MaxLogSizeBytes As Integer
    '  Get
    '    Try
    '      ' <Modified by: Ernie at 3/24/2014-4:13:47 PM on machine: ERNIE-THINK>
    '      ' Return My.Application.Log.DefaultFileLogWriter.MaxFileSize
    '      Return GetMaxLogFilesSizeBytes()
    '      ' </Modified by: Ernie at 3/24/2014-4:13:47 PM on machine: ERNIE-THINK>
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    '  Set(value As Integer)
    '    Try
    '      ' <Modified by: Ernie at 3/24/2014-4:19:19 PM on machine: ERNIE-THINK>
    '      ' My.Application.Log.DefaultFileLogWriter.MaxFileSize = value
    '      SetMaxLogFileSizeBytes(value)
    '      ' </Modified by: Ernie at 3/24/2014-4:19:19 PM on machine: ERNIE-THINK>
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

    'Public Shared Property MaxLogSizeMegaBytes As Integer
    '  Get
    '    Try
    '      Return (My.Application.Log.DefaultFileLogWriter.MaxFileSize / 1024 / 1024)
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    '  Set(value As Integer)
    '    Try
    '      My.Application.Log.DefaultFileLogWriter.MaxFileSize = (value * 1024 * 1024)
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

    Public Shared ReadOnly Property DefaultMaxLogSizeBytes As Integer
      Get
        Return MAX_LOG_SIZE
      End Get
    End Property

    Public Shared ReadOnly Property DefaultMaxLogSizeMegaBytes As Integer
      Get
        Return MAX_LOG_SIZE / 1024 / 1024
      End Get
    End Property

#End Region

#Region "Public Methods"

#Region "Initialization Methods"

    ''' <summary>
    ''' Sets up the application log as a daily log in a 'Logs' subdirectory of the current application path
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile()
      Try

        InitializeApplicationLogFile(Assembly.GetExecutingAssembly.GetName.Name)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Sets up the application log as a daily log in a 'Logs' subdirectory of the current application path
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile(lpAsJson As Boolean)
      Try

        InitializeApplicationLogFile(Assembly.GetExecutingAssembly.GetName.Name, lpAsJson)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Sets up the application log as a daily log in a 'Logs' subdirectory of the current application path
    ''' </summary>
    ''' <param name="lpMaxLogSize">The maximum size in bytes before the log rolls over to a new file.</param>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile(ByVal lpMaxLogSize As Integer)
      Try

        InitializeApplicationLogFile(Assembly.GetExecutingAssembly.GetName.Name, lpMaxLogSize)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Sets up the application log as a daily log in a 'Logs' subdirectory of the current application path
    ''' </summary>
    ''' <param name="lpMaxLogSize">The maximum size in bytes before the log rolls over to a new file.</param>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile(ByVal lpMaxLogSize As Integer, lpAsJson As Boolean)
      Try

        InitializeApplicationLogFile(Assembly.GetExecutingAssembly.GetName.Name, lpMaxLogSize, lpAsJson)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Sets up the application log as a daily log in a 'Logs' subdirectory of the current application path
    ''' </summary>
    ''' <param name="lpApplicationName">The name of the application</param>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile(ByVal lpApplicationName As String, Optional lpAsJson As Boolean = False)
      '  As Gurock.SmartInspect.Session
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpApplicationName)
#Else
        If lpApplicationName Is Nothing Then
          Throw New ArgumentNullException
        End If
#End If

        InitializeApplicationLogFile(lpApplicationName, LogDirectory, lpAsJson)

      Catch AppLogInitEx As Exceptions.ApplicationLogNotInitializedException
        If AppLogInitEx.InnerException IsNot Nothing AndAlso
         TypeOf (AppLogInitEx.InnerException) Is FileLoadException Then
          ' Try one more time, sometimes we get a FileLoadException like
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          ' Cannot resolve dependency to assembly '~~~~~~~~~~~~~~~' 
          ' because it has not been preloaded. When using the ReflectionOnly APIs, 
          ' dependent assemblies must be pre-loaded or loaded on demand through 
          ' the ReflectionOnlyAssemblyResolve event.
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '
          ' This can often be resolved by making an additional pass at it.
          ' It appears that it might be a timing issue
          Try
            InitializeApplicationLogFile(lpApplicationName, LogDirectory, lpAsJson)
          Catch ex As Exception
            ' Forget it, 
            Throw
          End Try
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Sets up the application log as a daily log in a 'Logs' subdirectory of the current application path
    ''' </summary>
    ''' <param name="lpApplicationName">The name of the application</param>
    ''' <param name="lpMaxLogSize">The maximum size in bytes before the log rolls over to a new file.</param>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile(ByVal lpApplicationName As String, ByVal lpMaxLogSize As Integer, Optional lpAsJson As Boolean = False)

      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpApplicationName)
#Else
        If lpApplicationName Is Nothing Then
          Throw New ArgumentNullException
        End If
#End If

        InitializeApplicationLogFile(lpApplicationName, LogDirectory, lpMaxLogSize, lpAsJson)

      Catch AppLogInitEx As Exceptions.ApplicationLogNotInitializedException
        If AppLogInitEx.InnerException IsNot Nothing AndAlso
         TypeOf (AppLogInitEx.InnerException) Is FileLoadException Then
          ' Try one more time, sometimes we get a FileLoadException like
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          ' Cannot resolve dependency to assembly '~~~~~~~~~~~~~~~' 
          ' because it has not been preloaded. When using the ReflectionOnly APIs, 
          ' dependent assemblies must be pre-loaded or loaded on demand through 
          ' the ReflectionOnlyAssemblyResolve event.
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '
          ' This can often be resolved by making an additional pass at it.
          ' It appears that it might be a timing issue
          Try
            InitializeApplicationLogFile(lpApplicationName, LogDirectory)
          Catch ex As Exception
            ' Forget it, 
            Throw
          End Try
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Sets up the application log as a daily log in the specified directory path
    ''' </summary>
    ''' <param name="lpApplicationName">The name of the application</param>
    ''' <param name="lpLogDirectory">The directory to write the log file to</param>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile(ByVal lpApplicationName As String,
                                                      ByVal lpLogDirectory As String, Optional lpAsJson As Boolean = False)
      Try
        'InitializeApplicationLogFile(lpApplicationName, lpLogDirectory, MaxLogSizeBytes, lpAsJson)
        InitializeApplicationLogFile(lpApplicationName, lpLogDirectory, 10000000, lpAsJson)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Sets up the application log as a daily log in the specified directory path
    ''' </summary>
    ''' <param name="lpApplicationName">The name of the application</param>
    ''' <param name="lpLogDirectory">The directory to write the log file to</param>
    ''' <param name="lpMaxLogSize">The maximum size in bytes before the log rolls over to a new file.</param>
    ''' <remarks></remarks>
    Public Shared Sub InitializeApplicationLogFile(ByVal lpApplicationName As String,
                                                      ByVal lpLogDirectory As String, ByVal lpMaxLogSize As Integer, Optional lpAsJson As Boolean = False)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpApplicationName)
        ArgumentNullException.ThrowIfNull(lpLogDirectory)
#Else
        If lpApplicationName Is Nothing Then
          Throw New ArgumentNullException(nameof(lpApplicationName))
        End If
        If lpLogDirectory Is Nothing Then
          Throw New ArgumentNullException(nameof(lpLogDirectory))
        End If
#End If


        'ConfigureNLogLogFile(lpApplicationName, lpLogDirectory)

        If Not System.IO.Directory.Exists(lpLogDirectory) Then
          System.IO.Directory.CreateDirectory(lpLogDirectory)
        End If
        ''With My.Application.Log.DefaultFileLogWriter
        ''  .Location = Logging.LogFileLocation.Custom
        ''  .CustomLocation = lpLogDirectory
        ''  .AutoFlush = True
        ''End With

        ''My.Application.Log.DefaultFileLogWriter.LogFileCreationSchedule = Logging.LogFileCreationScheduleOption.Daily
        ''My.Application.Log.DefaultFileLogWriter.MaxFileSize = lpMaxLogSize ' = 10485760 ' Ten MB
        ''My.Application.Log.DefaultFileLogWriter.DiskSpaceExhaustedBehavior =
        ''  Logging.DiskSpaceExhaustedOption.ThrowException
        ''My.Application.Log.DefaultFileLogWriter.BaseFileName = _
        ''  String.Format("{0}_{1}", lpApplicationName, Environment.MachineName)
        ''My.Application.Log.DefaultFileLogWriter.BaseFileName = GetCurrentBaseLogFile(lpApplicationName)
        ''My.Application.Log.DefaultFileLogWriter.TraceOutputOptions = TraceOptions.DateTime Or TraceOptions.ThreadId
        ''Trace.AutoFlush = True

        'Dim lobjAssemblyVersionInfo As New AssemblyVersionInfo

        'Try
        '  lobjAssemblyVersionInfo = AssemblyVersionInfo.Create(My.Application.Info)
        'Catch fle As FileLoadException
        '  ' Try one more time
        '  Try
        '    lobjAssemblyVersionInfo = AssemblyVersionInfo.Create(My.Application.Info)
        '  Catch ex As Exception
        '    Throw New ApplicationLogNotInitializedException( _
        '      String.Format("Unable to resolve assembly version information for {0}.", lpApplicationName), _
        '      lpApplicationName, lpLogDirectory, ex)
        '  End Try
        'End Try

        'Dim lstrMessage As String = String.Format("{0} '{1}' Loading under user '{2}', the current locale is '{3}'", _
        '                                            lpApplicationName, _
        '                                            lobjAssemblyVersionInfo.FileVersion.ToString, _
        '                                            WindowsIdentity.GetCurrent.Name, _
        '                                            Globalization.CultureInfo.CurrentCulture.Name)


        'Dim lstrBaseFileName = String.Format("{0}_{1}", lpApplicationName, Environment.MachineName)


        Select Case ConnectionSettings.Instance.LoggingFramework
          Case LoggingFramework.Serilog
            InitializeSerilog(lpApplicationName, lpMaxLogSize, lpAsJson)
          Case Else
            ConfigureLog4Net()
        End Select

        'Dim lstrBaseFileName As String = GetCurrentBaseLogFile(lpApplicationName, lpAsJson)

        ''Log.Logger = New LoggerConfiguration().WriteTo.File(lstrBaseFileName,,,, lpMaxLogSize).CreateLogger()r
        'Log.Logger = New LoggerConfiguration().WriteTo.File(New JsonFormatter(), lstrBaseFileName, LogEventLevel.Debug, lpMaxLogSize,,,, New TimeSpan(0, 0, 30), RollingInterval.Day, lpMaxLogSize).CreateLogger()

        'Dim lstrMessage As String = String.Format("{0} '{1}' Loading under user '{2}', the current locale is '{3}'",
        '                                        lpApplicationName,
        '                                        Assembly.GetExecutingAssembly.GetName().Version.ToString(),
        '                                        Environment.UserName,
        '                                        Globalization.CultureInfo.CurrentCulture.Name)

        ''My.Application.Log.WriteEntry(lstrMessage, TraceEventType.Information, OPEN_LOG_EVENT_ID)
        'Log.Information(lstrMessage)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Private Shared Sub ConfigureNLogLogFile(ByVal lpApplicationName As String, ByVal lpLogDirectory As String)
    '  Try

    '    If lpApplicationName Is Nothing Then
    '      Throw New ArgumentNullException("lpApplicationName")
    '    End If

    '    If lpLogDirectory Is Nothing Then
    '      Throw New ArgumentNullException("lpLogDirectory")
    '    End If

    '    If Not IO.Directory.Exists(lpLogDirectory) Then
    '      IO.Directory.CreateDirectory(lpLogDirectory)
    '    End If

    '    ' Step 1. Create configuration object 
    '    Dim lobjLoggingConfiguration As New LoggingConfiguration

    '    ' Step 2. Create targets and add them to the configuration 
    '    Dim lobjLogFileTarget As New FileTarget
    '    lobjLoggingConfiguration.AddTarget("LogFile", lobjLogFileTarget)

    '    ' Step 3. Set target properties 
    '    lobjLogFileTarget.FileName = Helper.CleanPath(String.Format("{0}\{1}_{2}", lpLogDirectory, lpApplicationName, Environment.MachineName))
    '    lobjLogFileTarget.Layout = "${message}"

    '    ' Step 4. Define rules
    '    Dim lobjFileLoggingRule As New LoggingRule("*", LogLevel.Debug, lobjLogFileTarget)
    '    lobjLoggingConfiguration.LoggingRules.Add(lobjFileLoggingRule)

    '    ' Step 5. Activate the configuration
    '    LogManager.Configuration = lobjLoggingConfiguration

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "Close Methods"

    ''' <summary>
    ''' Closes out the session in the applicationlog
    ''' </summary>
    ''' <remarks></remarks>
    'Public Shared Sub CloseApplicationLogFile()
    '  Try
    '    CloseApplicationLogFile(Assembly.GetExecutingAssembly.GetName.Name)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ''' <summary>
    ''' Closes out the session in the applicationlog
    ''' </summary>
    ''' <param name="lpApplicationName">The name of the application</param>
    ''' <remarks></remarks>
    'Public Shared Sub CloseApplicationLogFile(ByVal lpApplicationName As String)
    '  Try

    '    If lpApplicationName Is Nothing Then
    '      Throw New ArgumentNullException("lpApplicationName")
    '    End If

    '    ' Dim lobjAssemblyVersionInfo As New AssemblyVersionInfo(My.Application.Info)


    '    'My.Application.Log.WriteEntry(String.Format("{0} '{1}' Unloading under user '{2}'", _
    '    '                                      lpApplicationName, _
    '    '                                      lobjAssemblyVersionInfo.FileVersion.ToString, _
    '    '                                      WindowsIdentity.GetCurrent.Name), _
    '    '                                      TraceEventType.Information, CLOSE_LOG_EVENT_ID)

    '    My.Application.Log.WriteEntry(String.Format("{0} '{1}' Unloading under user '{2}'",
    '                                                lpApplicationName,
    '                                                My.Application.Info.Version.ToString,
    '                                                WindowsIdentity.GetCurrent.Name),
    '                                  TraceEventType.Information, CLOSE_LOG_EVENT_ID)

    '    ' Write a blank enry to separate application sessions
    '    My.Application.Log.DefaultFileLogWriter.WriteLine(ControlChars.CrLf)

    '    Trace.Flush()

    '    ApplicationLogging.WriteLogEntry(String.Format("{0} '{1}' Unloading under user '{2}'",
    '                                         lpApplicationName,
    '                                         My.Application.Info.Version.ToString,
    '                                         WindowsIdentity.GetCurrent.Name), TraceEventType.Information)

    '    Try
    '      'SiAuto.Main.LeaveProcess(lpApplicationName)
    '    Catch ex As Exception

    '    End Try


    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "Warning Logging Methods"

    Shared Sub LogInformation(ByVal lpTitleFormat As String, ParamArray lpArgs As Object())
      Try
        LogInformation(String.Format(lpTitleFormat, lpArgs))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Shared Sub LogInformation(ByVal lpMessage As String, Optional lpEchoToConsole As Boolean = False)
      Try
        WriteLogEntry(lpMessage, TraceEventType.Information)
        If lpEchoToConsole Then
          Console.WriteLine(lpMessage)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Shared Sub LogInformation(ByVal lpMessage As String, ByVal lpCallingMethod As Reflection.MethodBase)
      Try
        WriteLogEntry(String.Format("{0}: {1}::{2}", lpMessage, lpCallingMethod.DeclaringType.Name, lpCallingMethod.Name), TraceEventType.Information)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Shared Sub LogInformation(ByVal lpCallingMethod As Reflection.MethodBase, ByVal lpTitleFormat As String, ParamArray lpArgs As Object())
      Try
        Dim lstrWarningMessage As String = String.Format(lpTitleFormat, lpArgs)
        'WriteLogEntry(String.Format("{0}: {0}::{1}", lstrWarningMessage, lpCallingMethod.DeclaringType.Name, lpCallingMethod.Name))
        LogInformation(lstrWarningMessage, lpCallingMethod)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub
#End Region

#Region "Warning Logging Methods"

    Shared Sub LogWarning(ByVal lpTitleFormat As String, ParamArray lpArgs As Object())
      Try
        LogWarning(String.Format(lpTitleFormat, lpArgs))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Shared Sub LogWarning(ByVal lpWarningMessage As String, Optional lpEchoToConsole As Boolean = False)
      Try
        WriteLogEntry(lpWarningMessage, TraceEventType.Warning)
        If lpEchoToConsole Then
          Console.WriteLine(lpWarningMessage)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Shared Sub LogWarning(ByVal lpWarningMessage As String, ByVal lpCallingMethod As Reflection.MethodBase)
      Try
        WriteLogEntry(String.Format("{0}: {1}::{2}", lpWarningMessage, lpCallingMethod.DeclaringType.Name, lpCallingMethod.Name), TraceEventType.Warning)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Shared Sub LogWarning(ByVal lpCallingMethod As Reflection.MethodBase, ByVal lpTitleFormat As String, ParamArray lpArgs As Object())
      Try
        Dim lstrWarningMessage As String = String.Format(lpTitleFormat, lpArgs)
        'WriteLogEntry(String.Format("{0}: {0}::{1}", lstrWarningMessage, lpCallingMethod.DeclaringType.Name, lpCallingMethod.Name))
        LogWarning(lstrWarningMessage, lpCallingMethod)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Exception Logging Methods"

    ''' <summary>
    ''' Writes exception information along with the method that the exception occured in to the application log.
    ''' </summary>
    ''' <param name="lpException">The exception to write.</param>
    ''' <param name="lpLocation">The method that threw the exception.</param>
    ''' <remarks>Passes 0 as a default error ID</remarks>
    Shared Sub LogException(ByVal lpException As Exception, ByVal lpLocation As String)

      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpException)
        ArgumentNullException.ThrowIfNull(lpLocation)
#Else
        If lpException Is Nothing Then
          Throw New ArgumentNullException(nameof(lpException))
        End If
        If lpLocation Is Nothing Then
          Throw New ArgumentNullException(nameof(lpLocation))
        End If
#End If


        LogException(lpException, lpLocation, 0)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Writes exception information along with the method that the exception occured in to the application log.
    ''' </summary>
    ''' <param name="lpException">The exception to write.</param>
    ''' <param name="lpLocation">The method that threw the exception.</param>
    ''' <param name="lpId">The ID to associate with the error in the application log.</param>
    ''' <remarks></remarks>
    Shared Sub LogException(ByVal lpException As Exception, ByVal lpLocation As String, ByVal lpId As Integer)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpException)
        ArgumentNullException.ThrowIfNull(lpLocation)
#Else
        If lpException Is Nothing Then
          Throw New ArgumentNullException(nameof(lpException))
        End If
        If lpLocation Is Nothing Then
          Throw New ArgumentNullException(nameof(lpLocation))
        End If
#End If

        '' For NLog
        'Dim lobjLogger As Logger = LogManager.GetLogger("ApplicationLogging")
        'lobjLogger.Error(String.Format("An Exception of type {0} occured in {1}: {2}", _
        '                               lpException.GetType.Name, lpLocation, lpException.Message))

        ExceptionTracker.Update(lpException, lpLocation)

        'If (My.Application Is Nothing) OrElse (My.Application.Log Is Nothing) Then
        '  ' If we do not have a valid object reference to the log then we can't write to it.
        '  Exit Sub
        'End If

        'My.Application.Log.WriteEntry(String.Format("An Exception of type {0} occured in {1}: {2}",
        '                                            lpException.GetType.Name, lpLocation, lpException.Message),
        '                              TraceEventType.Error, lpId)

        ApplicationLogging.WriteLogEntry(String.Format("An Exception of type {0} occured in {1}: {2}",
                                             lpException.GetType.Name, lpLocation, lpException.Message), TraceEventType.Error)

        Debug.WriteLine(String.Format("Exception Type: {0}", lpException.GetType.Name))

        '  For selected exception types write additional information to the log
        Select Case lpException.GetType.Name
          Case "InvalidPathException"
            ApplicationLogging.WriteLogEntry(String.Format("The path '{0}' is invalid.",
                                                      CType(lpException, Object).Path), TraceEventType.Error)
          Case "FileNotFoundException"
            ApplicationLogging.WriteLogEntry(String.Format("The file '{0}' was not found.",
                                                      CType(lpException, FileNotFoundException).FileName), TraceEventType.Error)
          Case "ReflectionTypeLoadException"
            Dim lintExceptionCounter As Integer
            For Each lobjException As Exception In
            CType(lpException, Reflection.ReflectionTypeLoadException).LoaderExceptions
              LogException(lobjException, String.Format("{0}~LoaderException({1})", lpLocation, lintExceptionCounter),
                         lpId)
              lintExceptionCounter += 1
            Next
        End Select

        '  If there is an inner exception make sure that it gets logged as well
        '  Often there is important information buried in the inner exception
        If lpException.InnerException IsNot Nothing Then
          LogException(lpException.InnerException, String.Format("{0}~InnerException", lpLocation), lpId)
        End If
      Catch InvalidOpEx As InvalidOperationException
        If _
        InvalidOpEx.Message =
        "Unable to write to log file because writing to it would cause it to exceed the MaxFileSize value." Then
          ' We need to create a new log file
          IncrementBaseLogFileName()
          LogException(lpException, lpLocation)
        End If
      Catch ex As Exception
        ' If we have a problem here then there is not anywhere else to go
        ' It is possible that this could be called by a non .NET application 
        ' or some other scenario in which My.Application.Log
        ' does not resolve to a valid object reference
        Exit Sub
      End Try
    End Sub

    ''' <summary>
    ''' Writes exception information along with the method that the exception occured in to the application log.
    ''' </summary>
    ''' <param name="lpException">The exception to write.</param>
    ''' <param name="lpCallingMethod">The method that threw the exception.</param>
    ''' <remarks>Passes 0 as a default error ID</remarks>
    Shared Sub LogException(ByVal lpException As Exception, ByVal lpCallingMethod As Reflection.MethodBase)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpException)
        ArgumentNullException.ThrowIfNull(lpCallingMethod)
#Else
        If lpException Is Nothing Then
          Throw New ArgumentNullException(nameof(lpException))
        End If
        If lpCallingMethod Is Nothing Then
          Throw New ArgumentNullException(nameof(lpCallingMethod))
        End If
#End If

        LogException(lpException, lpCallingMethod, 0)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Writes exception information along with the method that the exception occured in to the application log.
    ''' </summary>
    ''' <param name="lpException">The exception to write.</param>
    ''' <param name="lpCallingMethod">The method that threw the exception.</param>
    ''' <param name="lpId">The ID to associate with the error in the application log.</param>
    ''' <remarks></remarks>
    Shared Sub LogException(ByVal lpException As Exception, ByVal lpCallingMethod As Reflection.MethodBase,
                          ByVal lpId As Integer)
      '      Helper.LogException(ex, String.Format("{0}::{1}", Reflection.MethodBase.GetCurrentMethod.DeclaringType.Name, Reflection.MethodBase.GetCurrentMethod.Name))

#If NET8_0_OR_GREATER Then
      ArgumentNullException.ThrowIfNull(lpCallingMethod)
#Else
        If lpCallingMethod Is Nothing Then
          Throw New ArgumentNullException(nameof(lpCallingMethod))
        End If
#End If

      LogException(lpException, $"{lpCallingMethod.DeclaringType.Name}::{lpCallingMethod.Name}-{lpId}")
    End Sub

    ''' <summary>
    ''' Writes exception information along with the method that the exception occured in to the application log.
    ''' </summary>
    ''' <param name="lpException">The exception to write.</param>
    ''' <param name="lpCallingMethod">The method that threw the exception.</param>
    ''' <param name="lpParameterValues"></param>
    ''' <remarks>Deprecated - Call Method by same name from ApplicationLogging instead of Helper</remarks>
    Shared Sub LogException(ByVal lpException As Exception,
                          ByVal lpCallingMethod As Reflection.MethodBase,
                          ByVal ParamArray lpParameterValues() As Object)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpException)
        ArgumentNullException.ThrowIfNull(lpCallingMethod)
        ArgumentNullException.ThrowIfNull(lpParameterValues)
#Else
        If lpException Is Nothing Then
          Throw New ArgumentNullException(nameof(lpException))
        End If
        If lpCallingMethod Is Nothing Then
          Throw New ArgumentNullException(nameof(lpCallingMethod))
        End If
        If lpParameterValues Is Nothing Then
          Throw New ArgumentNullException(nameof(lpParameterValues))
        End If
#End If

        LogException(lpException, CreateMethodSignatureLabel(lpCallingMethod, lpParameterValues))
      Catch ex As Exception
        LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

    ''' <summary>
    ''' Writes exception information along with the method that the exception occured in to the application log.
    ''' </summary>
    ''' <param name="lpException">The exception to write.</param>
    ''' <param name="lpCallingMethod">The method that threw the exception.</param>
    ''' <param name="lpId">The ID to associate with the error in the application log.</param>
    ''' <param name="lpParameterValues"></param>
    ''' <remarks>Deprecated - Call Method by same name from ApplicationLogging instead of Helper</remarks>
    Shared Sub LogException(ByVal lpException As Exception,
                          ByVal lpCallingMethod As Reflection.MethodBase,
                          ByVal lpId As Integer,
                          ByVal ParamArray lpParameterValues() As Object)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpCallingMethod)
#Else
          If lpCallingMethod Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpCallingMethod))
          End If
#End If

        LogException(lpException, CreateMethodSignatureLabel(lpCallingMethod, lpParameterValues), lpId)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

#End Region

#Region "Write Log Entry"

    ''' <summary>
    ''' Writes an entry to the application log
    ''' </summary>
    ''' <param name="lpMessage">The message to write to the log</param>
    ''' <remarks>Event Type defaults to Information</remarks>
    Public Shared Sub WriteLogEntry(ByVal lpMessage As String)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpMessage)
#Else
          If lpMessage Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpMessage))
          End If
#End If
        WriteLogEntry(lpMessage, TraceEventType.Information, 0)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Writes an entry to the application log
    ''' </summary>
    ''' <param name="lpMessage">The message to write to the log</param>
    ''' <param name="lpEventType">The event severity</param>
    ''' <remarks>Event Id defaults to 0</remarks>
    Public Shared Sub WriteLogEntry(ByVal lpMessage As String, ByVal lpEventType As TraceEventType)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpMessage)
#Else
          If lpMessage Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpMessage))
          End If
#End If

        WriteLogEntry(lpMessage, lpEventType, 0)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Writes an entry to the application log
    ''' </summary>
    ''' <param name="lpMessage">The message to write to the log</param>
    ''' <param name="lpEventType">The event severity</param>
    ''' <param name="lpId">The Id associated with the event</param>
    ''' <remarks></remarks>
    Public Shared Sub WriteLogEntry(ByVal lpMessage As String, ByVal lpEventType As TraceEventType,
                                  ByVal lpId As Integer)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpMessage)
#Else
          If lpMessage Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpMessage))
          End If
#End If

        WriteLogEntry(lpMessage, String.Empty, lpEventType, lpId)
        'WriteSerilogEntry(lpMessage, lpEventType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub WriteLogEntry(ByVal lpMessage As String,
                                  ByVal lpCallingMethod As Reflection.MethodBase,
                                  ByVal lpEventType As TraceEventType,
                                  ByVal lpId As Integer)
      Try
        WriteLogEntry(lpMessage, CreateMethodSignatureLabel(lpCallingMethod), lpEventType, lpId)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub WriteLogEntry(ByVal lpMessage As String,
                                  ByVal lpCallingMethod As Reflection.MethodBase,
                                  ByVal lpEventType As TraceEventType,
                                  ByVal lpId As Integer,
                                  ByVal ParamArray lpParameterValues() As Object)
      Try
        WriteLogEntry(lpMessage, CreateMethodSignatureLabel(lpCallingMethod, lpParameterValues), lpEventType, lpId)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Shared Function TranslateTraceLevel(lpEventType As TraceEventType) As LogEventLevel
      Select Case lpEventType
        Case TraceEventType.Critical
          Return LogEventLevel.Fatal
        Case TraceEventType.Error
          Return LogEventLevel.Error
        Case TraceEventType.Warning
          Return LogEventLevel.Warning
        Case TraceEventType.Information
          Return LogEventLevel.Information
        Case TraceEventType.Verbose
          Return LogEventLevel.Verbose
        Case Else
          Return LogEventLevel.Information
      End Select
    End Function



    Public Shared Sub WriteLogEntry(ByVal lpMessage As String,
                                  ByVal lpLocation As String,
                                  ByVal lpEventType As TraceEventType,
                                  ByVal lpId As Integer)
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpMessage)
#Else
          If lpMessage Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpMessage))
          End If
#End If

        Dim lstrMessage As String
        If Not String.IsNullOrEmpty(lpLocation) Then
          lstrMessage = $"{lpMessage}: {lpLocation}"
        Else
          lstrMessage = lpMessage
        End If

        Select Case ConnectionSettings.Instance.LoggingFramework
          Case LoggingFramework.Serilog
            WriteSerilogEntry(lstrMessage, lpEventType)
          Case LoggingFramework.Log4Net
            WriteLog4NetEntry(lstrMessage, lpEventType)
          Case Else
            WriteLog4NetEntry(lstrMessage, lpEventType)
        End Select

        'If Not String.IsNullOrEmpty(lpLocation) Then
        '  'My.Application.Log.WriteEntry(String.Format("{0}: {1}", lpMessage, lpLocation),
        '  '                              lpEventType, lpId)
        '  WriteSerilogEntry(String.Format("{0}: {1}", lpMessage, lpLocation), lpEventType)
        '  'SiAuto.Main.LogMessage(String.Format("{0}: {1}", lpMessage, lpLocation))
        'Else
        '  'My.Application.Log.WriteEntry(lpMessage, lpEventType, lpId)
        '  'My.Application.Log.WriteEntry(String.Format("{0}", lpMessage), lpEventType, lpId)
        '  WriteSerilogEntry(lpMessage, lpEventType)
        '  'SiAuto.Main.LogMessage(lpMessage)
        '  'SiAuto.Main.LogString(Level.Verbose, lpLocation, lpMessage)
        'End If

        ''If Not String.IsNullOrEmpty(lpLocation) Then
        ''  Dim lstrLogMessage As String = $"{lpLocation}: {lpMessage}"
        ''  Select Case lpEventType

        ''    Case TraceEventType.Information
        ''      ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Information)
        ''    Case TraceEventType.Error
        ''      ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Error)
        ''    Case TraceEventType.Warning
        ''      ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Warning)
        ''    Case TraceEventType.Verbose
        ''      ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Verbose)
        ''    Case Else
        ''      ApplicationLogging.WriteLogEntry(lstrLogMessage, TraceEventType.Information)
        ''  End Select
        ''Else
        ''  Select Case lpEventType
        ''    Case TraceEventType.Information
        ''      ApplicationLogging.WriteLogEntry(lpMessage, TraceEventType.Information)
        ''    Case TraceEventType.Warning
        ''      ApplicationLogging.WriteLogEntry(lpMessage, TraceEventType.Warning)
        ''    Case TraceEventType.Verbose
        ''      ApplicationLogging.WriteLogEntry(lpMessage, TraceEventType.Verbose)
        ''    Case TraceEventType.Error
        ''      ApplicationLogging.WriteLogEntry(lpMessage, TraceEventType.Error)
        ''    Case Else
        ''      ApplicationLogging.WriteLogEntry(lpMessage, TraceEventType.Information)
        ''  End Select
        ''End If


      Catch InvalidOpEx As InvalidOperationException
        If _
        InvalidOpEx.Message =
        "Unable to write to log file because writing to it would cause it to exceed the MaxFileSize value." Then
          ' We need to create a new log file
          IncrementBaseLogFileName()
          WriteLogEntry(lpMessage, lpLocation, lpEventType, lpId)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    '#Region "Write To Event Log"

    '    Public Shared Sub WriteToEventLog(ByVal lpEntry As String)
    '      Try
    '        'WriteToEventLog(lpEntry, Assembly.GetExecutingAssembly.GetName.Name, EventLogEntryType.Information, "Application", 0)
    '        WriteToEventLog(lpEntry, Assembly.GetExecutingAssembly.GetName.Name, EventLogEntryType.Information, "Application", 0)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        'Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '    Public Shared Sub WriteToEventLog(ByVal lpEntry As String,
    '                                    ByVal lpEventType As EventLogEntryType)
    '      Try
    '        WriteToEventLog(lpEntry, Assembly.GetExecutingAssembly.GetName.Name, lpEventType, "Application", 0)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        'Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '    Public Shared Sub WriteToEventLog(ByVal lpEntry As String,
    '                                    ByVal lpEventType As EventLogEntryType,
    '                                    ByVal lpId As Integer)
    '      Try
    '        WriteToEventLog(lpEntry, Assembly.GetExecutingAssembly.GetName.Name, lpEventType, "Application", lpId)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        'Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '    Public Shared Sub WriteToEventLog(ByVal lpEntry As String,
    '                                    ByVal lpEventType As EventLogEntryType,
    '                                    ByVal lpLogName As String,
    '                                    ByVal lpId As Integer)
    '      Try
    '        WriteToEventLog(lpEntry, Assembly.GetExecutingAssembly.GetName.Name, lpEventType, lpLogName, lpId)
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '        'Re-throw the exception to the caller
    '        Throw
    '      End Try
    '    End Sub

    '    Public Shared Sub WriteToEventLog(ByVal lpEntry As String,
    '                                    ByVal lpApplicationName As String,
    '                                    ByVal lpEventType As EventLogEntryType,
    '                                    ByVal lpLogName As String,
    '                                    ByVal lpId As Integer)

    '      '*************************************************************
    '      'PURPOSE: Write Entry to Event Log using VB.NET
    '      'PARAMETERS: Entry - Value to Write
    '      '            AppName - Name of Client Application. Needed 
    '      '              because before writing to event log, you must 
    '      '              have a named EventLog source. 
    '      '            EventType - Entry Type, from EventLogEntryType 
    '      '              Structure e.g., EventLogEntryType.Warning, 
    '      '              EventLogEntryType.Error
    '      '            LogName: Name of Log (System, Application; 
    '      '              Security is read-only) If you 
    '      '              specify a non-existent log, the log will be
    '      '              created

    '      'RETURNS:   True if successful, false if not

    '      'EXAMPLES: 
    '      '1. Simple Example, Accepting All Defaults
    '      '    WriteToEventLog "Hello Event Log"

    '      '2.  Specify EventSource, EventType, and LogName
    '      '    WriteToEventLog("Danger, Danger, Danger", "MyVbApp", _
    '      '                      EventLogEntryType.Warning, "System")
    '      '
    '      'NOTE:     EventSources are tightly tied to their log. 
    '      '          So don't use the same source name for different 
    '      '          logs, and vice versa
    '      '******************************************************

    '      Dim objEventLog As New EventLog()

    '      Try
    '        'Register the App as an Event Source
    '        If Not EventLog.SourceExists(lpApplicationName) Then
    '          EventLog.CreateEventSource(lpApplicationName, lpLogName)
    '        End If

    '        objEventLog.Source = lpApplicationName

    '        'WriteEntry is overloaded; this is one
    '        'of 10 ways to call it
    '        objEventLog.WriteEntry(lpEntry, lpEventType, lpId)
    '      Catch Ex As Exception
    '        EventLog.WriteEntry(Assembly.GetExecutingAssembly.GetName.Name, Ex.Message)
    '      End Try
    '    End Sub

    '#End Region

#Region "Public Methods"

    'Public Shared Sub Flush()
    '  Try
    '    My.Application.Log.DefaultFileLogWriter.Flush()
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Shared Function CreateMethodSignatureLabel(ByVal lpCallingMethod As Reflection.MethodBase) As String
      Try
        Return CreateMethodSignatureLabel(lpCallingMethod, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return ""
      End Try
    End Function

    Public Shared Function CreateMethodSignatureLabel(ByVal lpCallingMethod As Reflection.MethodBase,
                                                    ByVal ParamArray lpParameterValues() As Object) As String
      Try
        Dim lstrOutputString As String

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpCallingMethod)
#Else
          If lpCallingMethod Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpCallingMethod))
          End If
#End If

        ' Build the first part of the output string
        lstrOutputString = String.Format("{0}::{1}", lpCallingMethod.DeclaringType.Name, lpCallingMethod.Name)

        If lpParameterValues IsNot Nothing Then

          If lpParameterValues.Length > 0 Then
            lstrOutputString &= "("
          End If

          For lintParameterCounter As Integer = 0 To lpParameterValues.Length - 1
            lstrOutputString &= String.Format("'{0}', ", lpParameterValues(lintParameterCounter).ToString)
          Next

          If lpParameterValues.Length > 0 Then
            lstrOutputString = lstrOutputString.Remove(lstrOutputString.Length - 2, 2)
            lstrOutputString &= ")"
          End If

        End If

        Return lstrOutputString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return ""
      End Try
    End Function

#End Region

#End Region

#Region "Private Methods"

    ''Private Shared Function GetLogFileName() As String
    ''  Try
    ''    If (My.Application.Log IsNot Nothing) AndAlso (My.Application.Log.DefaultFileLogWriter IsNot Nothing) Then
    ''      Return My.Application.Log.DefaultFileLogWriter.FullLogFileName
    ''    Else
    ''      Return String.Empty
    ''    End If
    ''  Catch ex As Exception
    ''    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    ''    ' Re-throw the exception to the caller
    ''    Throw
    ''  End Try
    ''End Function

    Private Shared Function GetCurrentBaseLogFile(lpApplicationName As String, Optional lpAsJson As Boolean = False) As String
      Try
        'Dim lstrCurrentLogFilePath As String = My.Application.Log.DefaultFileLogWriter.FullLogFileName
        'Dim lstrCurrentLogFileName As String = Path.GetFileName(lstrCurrentLogFilePath)
        'Dim lstrCurrentBaseLogFileName As String = My.Application.Log.DefaultFileLogWriter.BaseFileName
        ''Dim lobjDefaultLogWriter As Logging.FileLogTraceListener = My.Application.Log.DefaultFileLogWriter

        Dim lstrCurrentBaseLogFileName As String
        ''If lobjDefaultLogWriter IsNot Nothing Then
        ''  lstrCurrentBaseLogFileName = lobjDefaultLogWriter.BaseFileName
        ''Else
        ''  lstrCurrentBaseLogFileName = lpApplicationName
        ''End If

        If Not String.IsNullOrEmpty(lpApplicationName) Then
          lstrCurrentBaseLogFileName = lpApplicationName
        Else
          lstrCurrentBaseLogFileName = BaseLogFileName
        End If
        'Dim lstrCurrentLogFileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(lstrCurrentLogFilePath)
        ' Dim lstrCurrentDatePart As String = lstrCurrentLogFileNameWithoutExtension.Substring(lstrCurrentLogFileNameWithoutExtension.Length - 10)
        'Dim lstrCurrentDatePart As String =
        '      lstrCurrentLogFileNameWithoutExtension.Replace(lstrCurrentBaseLogFileName, String.Empty).TrimStart("-")
        'Dim lstrLogDirectory As String = Path.GetDirectoryName(lstrCurrentLogFilePath)
        Dim lstrLogDirectory As String = LogDirectory 'ConnectionSettings.Instance.LogFolderPath
        Dim lstrDefaultBaseLogFileName As String

        If lpAsJson Then
          lstrDefaultBaseLogFileName = String.Format("{0}_{1}_log.json", lpApplicationName, Environment.MachineName) 'Environment.MachineName)
        Else
          lstrDefaultBaseLogFileName = String.Format("{0}_{1}.log", lpApplicationName, Environment.MachineName)
        End If

        'Dim lstrTodaysLogFileName As String = String.Format("{0}-{1}", lstrDefaultBaseLogFileName, lstrCurrentDatePart)

        Dim lstrLogFilePath As String = Path.Combine(lstrLogDirectory, lstrDefaultBaseLogFileName)

        'Dim lstrTodaysLogFileNames As String() = Directory.GetFiles(lstrLogDirectory, String.Format("{0}*{1}.log",
        '                                                                                            lstrDefaultBaseLogFileName,
        '                                                                                            lstrCurrentDatePart))
        Dim lstrReturnValue As String = String.Empty

        'If lstrTodaysLogFileNames.Count > 1 Then
        '  Dim lstrLatestLogFileName As String = Path.GetFileNameWithoutExtension(lstrTodaysLogFileNames.Last)
        '  Return lstrLatestLogFileName.Remove(lstrLatestLogFileName.Length - 11)
        'Else
        '  If File.Exists(lstrCurrentLogFilePath) Then
        '    'Dim lobjFileInfo As New FileInfo(lstrCurrentLogFilePath)
        '    'If lobjFileInfo.Length >= MAX_LOG_SIZE Then
        '    '  retu()
        '    'Else
        '    '  Return lstrDefaultBaseLogFileName
        '    'End If
        '    Return lstrCurrentLogFilePath
        '  End If
        '  ' Return lstrCurrentLogFileName
        '  Return lstrDefaultBaseLogFileName
        'End If

        'If lstrTodaysLogFileNames.Count > 0 Then
        '  Dim lstrLatestLogFileName As String = Path.GetFileNameWithoutExtension(lstrTodaysLogFileNames.Last)
        '  lstrReturnValue = lstrLatestLogFileName.Remove(lstrLatestLogFileName.Length - 11)
        'Else
        '  lstrReturnValue = lstrDefaultBaseLogFileName
        'End If

        lstrReturnValue = lstrLogFilePath

        'My.Application.Log.DefaultFileLogWriter.BaseFileName = lstrReturnValue

        '' Sometimes it will create an empty default log file with the first call 
        '' to My.Application.Log.DefaultFileLogWriter.FullLogFileName
        '' 
        '' If it did we want to clean it up
        'If File.Exists(lstrCurrentLogFilePath) Then
        '  Dim lobjFileInfo As New FileInfo(lstrCurrentLogFilePath)
        '  If lobjFileInfo.Length < 10 Then
        '    ' This is a tiny file, it should be safe to go ahead and delete it.
        '    lobjFileInfo.Delete()
        '  End If
        'End If

        Return lstrReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Private Shared Function GetCurrentLogFile(lpApplicationName As String) As String
    '  Try
    '    'Dim lstrCurrentLogFileName As String = My.Application.Log.DefaultFileLogWriter.FullLogFileName
    '    'Dim lstrLogDirectory As String = Path.GetDirectoryName(lstrCurrentLogFileName)
    '    'Dim lstrCurrentBaseLogFileName As String = String.Format("{0}_{1}", lpApplicationName, Environment.MachineName)
    '    'Dim lstrTodaysLogFileNames As String() = Directory.GetFiles(lstrLogDirectory, String.Format("{0}*", _
    '    '                                                      lstrCurrentBaseLogFileName))

    '    'If lstrTodaysLogFileNames.Count > 1 Then
    '    '  Dim lstrLatestLogFileName As String = Path.GetFileNameWithoutExtension(lstrTodaysLogFileNames.Last)
    '    '  Return lstrLatestLogFileName.Remove(lstrLatestLogFileName.Length - 11)
    '    'Else
    '    '  Return lstrCurrentBaseLogFileName
    '    'End If

    '    Return Path.GetFileName(My.Application.Log.DefaultFileLogWriter.FullLogFileName)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Private Shared Function IncrementBaseLogFileName() As String
      Try
        'Dim lstrCurrentLogFileName As String = My.Application.Log.DefaultFileLogWriter.FullLogFileName
        'Dim lstrLogDirectory As String = Path.GetDirectoryName(lstrCurrentLogFileName)
        'Dim lstrCurrentBaseLogFileName As String = My.Application.Log.DefaultFileLogWriter.BaseFileName
        'Dim lstrTodaysLogFileNames As String() = Directory.GetFiles(lstrLogDirectory, String.Format("{0}*", _
        '                                                              lstrCurrentBaseLogFileName))

        'If lstrTodaysLogFileNames.Count > 1 Then
        '  'Dim lobjLogFileNames As New List(Of String)
        '  'For Each lstrLogFileName As String In lstrTodaysLogFileNames
        '  '  lobjLogFileNames.Add(lstrLogFileName)
        '  'Next
        '  'lobjLogFileNames.Sort()
        '  'lstrCurrentBaseLogFileName = lobjLogFileNames.Last
        '  Dim lstrLatestLogFileName As String = Path.GetFileNameWithoutExtension(lstrTodaysLogFileNames.Last)
        '  lstrCurrentBaseLogFileName = lstrLatestLogFileName.Remove(lstrLatestLogFileName.Length - 11)
        '  ' lstrCurrentBaseLogFileName = Path.GetFileNameWithoutExtension(lstrTodaysLogFileNames.Last)
        'End If

        ' <Modified by: Ernie at 3/24/2014-3:55:00 PM on machine: ERNIE-THINK>
        ' Dim lstrCurrentBaseLogFileName As String = GetCurrentBaseLogFile("Ecmg.Cts.UI.Wpf.JobManager")
        Dim lstrCurrentBaseLogFileName As String = GetCurrentBaseLogFile(Assembly.GetExecutingAssembly.GetName.Name)
        ' </Modified by: Ernie at 3/24/2014-3:55:00 PM on machine: ERNIE-THINK>

        Dim lstrIncrementedLogFileName As String


        If lstrCurrentBaseLogFileName.EndsWith(Environment.MachineName) Then
          'My.Application.Log.DefaultFileLogWriter.BaseFileName &= "_1"
          lstrIncrementedLogFileName = $"{lstrCurrentBaseLogFileName}_1"
        Else
          Dim lstrLastLetter As String = lstrCurrentBaseLogFileName.Substring(lstrCurrentBaseLogFileName.Length - 1)
          Dim lobjStringBuilder As New StringBuilder
          lobjStringBuilder.Append(lstrCurrentBaseLogFileName.AsSpan(0, lstrCurrentBaseLogFileName.Length - 1))
          lobjStringBuilder.Append(GetNextLetter(lstrLastLetter))
          'My.Application.Log.DefaultFileLogWriter.BaseFileName = lobjStringBuilder.ToString
          lstrIncrementedLogFileName = lobjStringBuilder.ToString
        End If

        'Return My.Application.Log.DefaultFileLogWriter.BaseFileName
        Return lstrIncrementedLogFileName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetNextLetter(lpCurrentLetter As String) As String
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpCurrentLetter)
#Else
          If lpCurrentLetter Is Nothing Then
            Throw New ArgumentNullException(NameOf(lpCurrentLetter))
          End If
#End If

        If lpCurrentLetter.Length > 1 Then
          Throw New InvalidOperationException("Only one character expected.")
        End If

        Dim lintASCII As Integer = Asc(lpCurrentLetter)

        Select Case lintASCII
          Case 65 To 89
            ' If we get A-Y return the next upper case letter
            Return Chr(lintASCII + 1).ToString
          Case 97 To 121
            ' If we get a-y return the next lower case letter
            Return Chr(lintASCII + 1).ToString
          Case 48 To 56
            ' If we get 0-8 return the next number
            Return Chr(lintASCII + 1).ToString
          Case 57
            ' If we get a 9 return a
            Return Chr(97).ToString
          Case 90
            ' If we get a Z return 0
            Return Chr(48).ToString
          Case 122
            ' If we get a z return 0
            Return Chr(48).ToString
          Case Else
            Throw _
              New InvalidOperationException($"Unexpected current letter '{lpCurrentLetter}', expected 0-9, a-z, A-Z.")
        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetMaxLogFilesSizeBytes() As Integer
      Try
        Dim lintMaxLogSize As Integer = DefaultMaxLogSizeBytes
        If Not String.IsNullOrEmpty(AppSettings(SETTING_MAX_LOG_SIZE)) Then
          Dim lstrMaxLogSize As String = AppSettings(SETTING_MAX_LOG_SIZE)
          If Integer.TryParse(lstrMaxLogSize, lintMaxLogSize) = False Then
            lintMaxLogSize = DefaultMaxLogSizeBytes
          End If
        Else
          ' Return My.Application.Log.DefaultFileLogWriter.MaxFileSize
          SetMaxLogFileSizeBytes(lintMaxLogSize)
        End If

        Return lintMaxLogSize

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Sub SetMaxLogFileSizeBytes(lpMaxBytes As Integer)
      Try
        If ApplicationHasMaxLogSizeKey() = False Then
          Dim lobjConfig As System.Configuration.Configuration =
              ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
          lobjConfig.AppSettings.Settings.Add(SETTING_MAX_LOG_SIZE, lpMaxBytes)
          lobjConfig.Save()
        End If
        AppSettings(SETTING_MAX_LOG_SIZE) = lpMaxBytes
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Shared Function ApplicationHasMaxLogSizeKey() As Boolean
      Try
        Dim lblnHasMaxLogSizeKey As Boolean = False
        For Each lstrKey As String In AppSettings.Keys
          If lstrKey = SETTING_MAX_LOG_SIZE Then
            lblnHasMaxLogSizeKey = True
            Exit For
          End If
        Next
        Return lblnHasMaxLogSizeKey
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region
  End Class

End Namespace
