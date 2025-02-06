'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Providers
Imports Documents.Serialization
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Configuration
  ''' <summary>
  ''' Singleton Class for accessing all connection settings.
  ''' </summary>
  ''' <remarks>Settings are persisted in a Settings.csf file</remarks>
  Public Class ConnectionSettings
    Implements IDisposable
    Implements ISerialize

#Region "Class Constants"

    Public Const CONNECTION_SETTINGS_FILE_EXTENSION As String = "csf"

#End Region

#Region "Class Variables"

    Private Shared mobjInstance As ConnectionSettings
    ' Private Shared mintReferenceCount As Integer
    Private mobjProviderPaths As New Collections.Specialized.StringCollection
    'Private mobjContentSourceConnectionStrings As New Collections.Specialized.StringCollection
    Private WithEvents MobjObservableContentSourceConnectionStrings As New ObservableCollection(Of String) _
    ' Collections.Specialized.StringCollection
    Private mobjContentSourceNames As List(Of String) = Nothing
    Private mobjContentSourceGroups As New ContentSourceGroups
    Private mobjProviderCatalog As ProviderCatalog
    'Private mstrUserName As String = Environment.UserName
    'Private mstrMachineName As String = Environment.MachineName
    Private mstrSettingsPath As String = String.Empty
    Private mstrProjectCatalogConnectionString As String = String.Empty
    Private mblnDisableNotifications As Boolean
    'Private menuLoggingLevel As Level = Level.Message
    Private mintNodeStatusRefreshInterval As Integer = 30
    Private mintJobStatusRefreshInterval As Integer = 30
    Private mintMaximumInMemoryDocumentMegabytes As Integer = 200
    Private mstrJobRunnerExecutionFilePath As String = String.Empty
    Private mintMaxBatchConcurrentThreads As Integer = 1
    Private mintMaxJobRunnerInstancesPerJob As Integer = 1
    Private mintMultipleJobRunnerInvocationDelaySeconds As Integer = 5
    Private mintJobRunnerConsoleFontSize As Integer = 0
    Private mblnJobRunnerDisplayProcesedMessage As Boolean = False
    Private mstrAtalasoftServiceTempFolderPath As String = String.Empty
    Private mstrLogFolderPath As String = String.Empty
    Private mstrFileMaticaServiceHost As String = "localhost"
    Private mintFileMaticaServicePort As Integer = 8080
    Private mstrFileMaticaServiceUri As String = "http://{0}:{1}/fmservice"
    'Private mobjLogSession As Gurock.SmartInspect.Session = Nothing

    ' Private Shared mobjConnectionSettings As ConnectionSettings = Nothing
    'Private mobjAssemblyCatalog As AssemblyCatalog
    'Private mobjAggregateCatalog As AggregateCatalog
    'Private mobjCompositionContainer As CompositionContainer
    Private mstrCurrentTheme As String = String.Empty

    Private mstrActiproLicenseeName As String = String.Empty
    Private mstrActiproLicenseKey As String = String.Empty
    Private menuLoggingFramework As ApplicationLogging.LoggingFramework = ApplicationLogging.LoggingFramework.Log4Net

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        'Compose()
        'InitializeLogSession(String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpXMLFilePath As String)
      Try
        Dim lstrErrorMessage As String = String.Empty

        'InitializeLogSession(lpXMLFilePath)
        'LogSession.EnterMethod(Level.Debug, Helper.GetMethodIdentifier("ConnectionSettings Constructor(lpXMLFilePath)"))

        'Compose()
        Dim lobjConnectionSettings As ConnectionSettings = Deserialize(lpXMLFilePath, lstrErrorMessage)

        If lstrErrorMessage.Length > 0 Then
          Throw _
            New ArgumentException(String.Format("Unable to create ConnectionSettings object from xml file: {0}",
                                                lstrErrorMessage))
        End If

        lobjConnectionSettings.mstrSettingsPath = lpXMLFilePath

        Helper.AssignObjectProperties(lobjConnectionSettings, Me)

        ' Since the provider catalog is read only we can't count on AssignObjectProperties to set it.
        mobjProviderCatalog = lobjConnectionSettings.ProviderCatalog


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      Finally
        'LogSession.LeaveMethod(Level.Debug, Helper.GetMethodIdentifier("ConnectionSettings Constructor(lpXMLFilePath)"))
      End Try
    End Sub

#End Region

#Region "Public Properties"

    <XmlIgnore()>
    Public ReadOnly Property ContentSourceNames As List(Of String)
      Get
        Try
          If mobjContentSourceNames Is Nothing OrElse mobjContentSourceNames.Count = 0 Then
            mobjContentSourceNames = GetContentSourceNames()
          End If
          Return mobjContentSourceNames
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlAttribute()>
    Public Property DisableNotifications As Boolean
      Get
        Try
          Return mblnDisableNotifications
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Boolean)
        Try
          mblnDisableNotifications = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlAttribute()>
    Public Property ActiproLicenseeName As String
      Get
        Try
          Return mstrActiproLicenseeName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrActiproLicenseeName = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlAttribute()>
    Public Property ActiproLicenseKey As String
      Get
        Try
          Return mstrActiproLicenseKey
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrActiproLicenseKey = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    '<XmlIgnore()>
    'Friend ReadOnly Property LogSession As Gurock.SmartInspect.Session
    '  Get
    '    Try
    '      Return mobjLogSession
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

    '<XmlAttribute()>
    'Public Property LoggingLevel As Level
    '  Get
    '    Try
    '      Return menuLoggingLevel
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    '  Set(value As Level)
    '    Try
    '      menuLoggingLevel = value
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Set
    'End Property

    Public ReadOnly Property ProjectCatalogConnectionStringInitialized As Boolean
      Get
        Try
          Return Not String.IsNullOrEmpty(mstrProjectCatalogConnectionString)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property LoggingFramework As ApplicationLogging.LoggingFramework
      Get
        Return menuLoggingFramework
      End Get
      Set(value As ApplicationLogging.LoggingFramework)
        menuLoggingFramework = value
      End Set
    End Property

    Public Property LogFolderPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrLogFolderPath) Then
            mstrLogFolderPath = "C:\CtsLogs"
          End If
          Return mstrLogFolderPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrLogFolderPath = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property FilematicaServiceHost As String
      Get
        Try
          Return mstrFileMaticaServiceHost
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrFileMaticaServiceHost = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property FileMaticaServicePort As Integer
      Get
        Try
          Return mintFileMaticaServicePort
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintFileMaticaServicePort = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property FileMaticaServiceUriTemplate As String
      Get
        Try
          Return mstrFileMaticaServiceUri
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrFileMaticaServiceUri = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property FileMaticaServiceUri As String
      Get
        Try
          Return String.Format(FileMaticaServiceUriTemplate, FilematicaServiceHost, FileMaticaServicePort)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ' <Import("ProjectCatalogConnectionString")> _
    Public Property ProjectCatalogConnectionString As String
      Get
        Try
          Return mstrProjectCatalogConnectionString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrProjectCatalogConnectionString = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property MaximumInMemoryDocumentMegabytes As Integer
      Get
        Try
          Return mintMaximumInMemoryDocumentMegabytes
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintMaximumInMemoryDocumentMegabytes = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property JobRunnerConsoleFontSize As Integer
      Get
        Try
          Return mintJobRunnerConsoleFontSize
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintJobRunnerConsoleFontSize = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property JobRunnerDisplayProcessedMessage As Boolean
      Get
        Try
          Return mblnJobRunnerDisplayProcesedMessage
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Boolean)
        Try
          mblnJobRunnerDisplayProcesedMessage = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property JobRunnerExecutionFilePath As String
      Get
        Try
          Return mstrJobRunnerExecutionFilePath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrJobRunnerExecutionFilePath = value
          InitializeJobRunnerExecutionFilePath()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property JobStatusRefreshInterval As Integer
      Get
        Try
          Return mintJobStatusRefreshInterval
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintJobStatusRefreshInterval = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property MaxBatchConcurrentThreads As Integer
      Get
        Try
          Return mintMaxBatchConcurrentThreads
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintMaxBatchConcurrentThreads = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property MaxJobRunnerInstancesPerJob As Integer
      Get
        Try
          Return mintMaxJobRunnerInstancesPerJob
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintMaxJobRunnerInstancesPerJob = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property MultipleJobRunnerInvocationDelaySeconds As Integer
      Get
        Try
          Return mintMultipleJobRunnerInvocationDelaySeconds
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintMultipleJobRunnerInvocationDelaySeconds = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property NodeStatusRefreshInterval As Integer
      Get
        Try
          Return mintNodeStatusRefreshInterval
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Integer)
        Try
          mintNodeStatusRefreshInterval = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property AtalasoftServiceTempFolderPath As String
      Get
        Try
          Return mstrAtalasoftServiceTempFolderPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrAtalasoftServiceTempFolderPath = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the collection of content source groups.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The groups are used to organize content sources in the tree view.</remarks>
    Public Property Groups As ContentSourceGroups
      Get
        If mobjContentSourceGroups.Count = 0 Then
          ' Initialize some groups as a test
          If ContentSourceNames.Count > 0 Then
            Dim lobjMainGroup As New ContentSourceList("Main")
            For Each lstrContentSource As String In ContentSourceNames
              lobjMainGroup.Add(lstrContentSource)
            Next
            mobjContentSourceGroups.Add(lobjMainGroup)
          End If
        End If
        Return mobjContentSourceGroups
      End Get
      Set(ByVal value As ContentSourceGroups)
        mobjContentSourceGroups = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or Sets the configured collection of ContentSource connection strings 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ContentSourceConnectionStrings() As ObservableCollection(Of String)
      Get
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return MobjObservableContentSourceConnectionStrings
      End Get
      Set(ByVal value As ObservableCollection(Of String))

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        MobjObservableContentSourceConnectionStrings = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or Sets the configured collection of provider paths 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ProviderPaths() As Collections.Specialized.StringCollection
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return mobjProviderPaths
      End Get
      Set(ByVal value As Collections.Specialized.StringCollection)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mobjProviderPaths = value
      End Set
    End Property

    ''' <summary>
    ''' Gets a ProviderCatalog for the listed ProviderPaths
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProviderCatalog() As Providers.ProviderCatalog
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If mobjProviderCatalog Is Nothing Then
            mobjProviderCatalog = GetProviderCatalog()
          End If

          Return mobjProviderCatalog

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property SettingsPath() As String
      Get
        Return mstrSettingsPath
      End Get
    End Property

    Public Property CurrentTheme As String
      Get
        Try
          If String.IsNullOrEmpty(mstrCurrentTheme) Then
            If _
              (Environment.OSVersion.Version.Major > 6) OrElse
              ((Environment.OSVersion.Version.Major = 6) AndAlso (Environment.OSVersion.Version.Minor >= 2)) Then
              mstrCurrentTheme = "MetroLight"
            Else
              mstrCurrentTheme = "OfficeBlack"
            End If
          End If
          Return mstrCurrentTheme
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrCurrentTheme = value
          If _
            Not Helper.IsDeserializationBasedCall() AndAlso Not Helper.CallStackContainsMethodName("GetCurrentSettings") _
            Then
            Save()
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property OperationExclusions() As ObservableCollection(Of String)


#End Region

#Region "Private Properties"

    Private ReadOnly Property IsDisposed() As Boolean
      Get
        Return disposedValue
      End Get
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Will get the current installed settings or create a 
    ''' new settings if one does not yet exist and return it.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCurrentSettings() As ConnectionSettings
      Try
        Return GetCurrentSettings(True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Will get the current installed settings or create a 
    ''' new settings if one does not yet exist and return it.
    ''' </summary>
    ''' <param name="lpForceRefresh">
    ''' Used to specify whether or not to 
    ''' always go back to the settings file 
    ''' or if we can use the cached settings if available.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetCurrentSettings(ByVal lpForceRefresh As Boolean) As ConnectionSettings
      Try

        If lpForceRefresh = True OrElse mobjInstance Is Nothing OrElse mobjInstance.IsDisposed = True Then
          Dim lstrSettingsFilePath As String

          'lstrSettingsFilePath = Helper.CleanPath(String.Format("{0}\Settings.csf",
          '                                                      FileHelper.Instance.ConfigPath))

          lstrSettingsFilePath = Helper.CleanPath(String.Format("{0}\Settings.csf",
                                                                Path.GetDirectoryName(Assembly.GetEntryAssembly.Location)))

          If File.Exists(lstrSettingsFilePath) Then
            mobjInstance = New ConnectionSettings(lstrSettingsFilePath)
          Else
            mobjInstance = Create(lstrSettingsFilePath)
          End If
        End If

        If mobjInstance.OperationExclusions IsNot Nothing Then
          If Not mobjInstance.OperationExclusions.Contains("TestOp") Then
            mobjInstance.OperationExclusions.Add("TestOp")
          End If
        End If

        'mobjInstance.InitializeLogSession()

        Return mobjInstance

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetCurrentSettingsText() As String
      Dim lobjReader As StreamReader = Nothing
      Try
        lobjReader = New StreamReader(GetCurrentSettingsPath)
        Return lobjReader.ReadToEnd
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lobjReader IsNot Nothing Then
          lobjReader.Close()
          lobjReader.Dispose()
        End If
      End Try
    End Function

    ''' <summary>
    ''' Returns the path found for current settings file
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetCurrentSettingsPath() As String
      Try

        Dim lstrSettingsFilePath As String

        lstrSettingsFilePath = String.Format("{0}\Settings.csf",
                                             FileHelper.Instance.ConfigPath)

        lstrSettingsFilePath = Helper.RemoveExtraBackSlashesFromFilePath(lstrSettingsFilePath)

        If File.Exists(lstrSettingsFilePath) Then
          Return lstrSettingsFilePath
        Else
          Return String.Empty
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Returns the path found for current settings file
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetExpectedSettingsPath() As String
      Try

        Dim lstrSettingsFilePath As String

        lstrSettingsFilePath = String.Format("{0}\Settings.csf",
                                             FileHelper.Instance.ConfigPath)


        Return lstrSettingsFilePath


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Attempts to add the provider path to the current settings file
    ''' </summary>
    ''' <param name="lpProviderPath">The fully qualified path to the provider</param>
    ''' <returns>Returns 0 if successful, 
    ''' -1 if the provider path is already in the ProviderPaths collection, or 
    ''' -2 if the provider path points to a provider utility file, or 
    ''' -3 if the provider path does not point to a valid provider</returns>
    ''' <remarks></remarks>
    Public Shared Function AddProviderToCurrentSettings(ByVal lpProviderPath As String) As Integer
      Try

        Dim lstrSettingsFilePath As String
        Dim lobjCurrentSettings As ConnectionSettings = GetCurrentSettings()

        ' First check to see if the provider path is already in the current settings
        For Each lstrProviderPath As String In lobjCurrentSettings.ProviderPaths
          If lstrProviderPath.Equals(lpProviderPath, StringComparison.CurrentCultureIgnoreCase) Then
            ' This path already exists, we will not add it again
            Return -1
          End If
        Next

        ' If we made it this far then we did not find the provider
#If NET8_0_OR_GREATER Then
        If lpProviderPath.Contains("utilities", StringComparison.CurrentCultureIgnoreCase) Then
          ' This is a utility file, we do not want to add it
          Return -2
        End If
#Else
        If lpProviderPath.ToLower.Contains("utilities") Then
          ' This is a utility file, we do not want to add it
          Return -2
        End If
#End If


        '' Second make sure the provider is a valid provider
        'Try
        '  Providers.ContentSource.GetProvider(lpProviderPath, , False)
        'Catch ex As Exception
        '  ' We were not able to create an instance of a provider 
        '  ' from the supplied path
        '  Return -3
        'End Try

        ' We will add it
        'lobjCurrentSettings.ProviderPaths.Add(lpProviderPath)

        If lpProviderPath.Contains("WindowsFileSystem") Then
          ' If the provider is the WindowsFileSystem Provider then
          ' Let's create a default ContentSource for it
          Dim lstrConnectionString As String
          lstrConnectionString =
            String.Format(
              "Name=Local File System;Provider=File System Provider;RootPath=Desktop;ExportPath=C:\CTS\FileSystem Exports\;ImportPath=C:\CTS\FileSystem Exports\;ProviderPath={0}",
              lpProviderPath)
          lobjCurrentSettings.ContentSourceConnectionStrings.Add(lstrConnectionString)
        End If

        ' Get the path to save the settings file to
        lstrSettingsFilePath = String.Format("{0}\Settings.csf",
                                             FileHelper.Instance.ConfigPath)
        ' Save the settings file
        lobjCurrentSettings.Serialize(lstrSettingsFilePath)

        Return 0

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Goes through all of the providers in the current Providers 
    ''' folder and attempts to add them to the current settings file
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub AddAllInstalledProvidersToCurrentSettings()
      Try

        ' Determine where the current provider folder is
        Dim lstrProvidersFolderPath As String
        Dim lstrCandidateProviderFiles() As String

        lstrProvidersFolderPath = FileHelper.Instance.ProviderPath

        lstrCandidateProviderFiles = Directory.GetFiles(lstrProvidersFolderPath, "Ecmg.Cts.Providers.*.dll")

        For Each lstrProviderCandidate As String In lstrCandidateProviderFiles
          If ContentSource.IsValidProvider(lstrProviderCandidate) = True Then
            Try
              AddProviderToCurrentSettings(lstrProviderCandidate)
            Catch ex As Exception
              ' Just try to log it and move on to the next one
              ' This is likely being called from InstallShield
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              Continue For
            End Try
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function GetConnectionString(ByVal lpSourceConnectionStrings As IEnumerable(Of String),
                                               ByVal lpContentSourceName As String) As String

      Try

        For Each lstrConnectionString As String In lpSourceConnectionStrings


#If NET8_0_OR_GREATER Then
          If (lstrConnectionString.ToLower.Contains("name=" & lpContentSourceName.ToLower, StringComparison.CurrentCultureIgnoreCase)) Then
            Return lstrConnectionString
          End If
#Else
          If (lstrConnectionString.ToLower.Contains("name=" & lpContentSourceName.ToLower)) Then
            Return lstrConnectionString
          End If
#End If

        Next

        Return String.Empty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetConnectionString(ByVal lpContentSourceName As String) As String

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return GetConnectionString(Me.ContentSourceConnectionStrings, lpContentSourceName)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetContentSourceConnectionStringIndex(ByVal lpContentSourceName As String) As Integer
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Dim i As Integer = 0
        If Me.ContentSourceConnectionStrings.Count > 0 Then
          For Each lstrConnectionString As String In Me.ContentSourceConnectionStrings


#If NET8_0_OR_GREATER Then
            If (lstrConnectionString.Contains("name=" & lpContentSourceName, StringComparison.CurrentCultureIgnoreCase)) Then
              Return i
            End If
#Else
            If (lstrConnectionString.ToLower.Contains("name=" & lpContentSourceName.ToLower)) Then
              Return i
            End If
#End If

            i += 1
          Next
        End If

        Return -1

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetContentSources() As Providers.ContentSources
      Try

        'LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Dim lobjContentSources As New Providers.ContentSources

        If Me.ContentSourceConnectionStrings.Count > 0 Then
          'LogSession.LogMessage("Getting ContentSources from ContentSourceConnectionStrings")
          For Each lstrConnectionString As String In Me.ContentSourceConnectionStrings
            lobjContentSources.Add(lstrConnectionString)
          Next
        End If

        'LogSession.LogMessage("Found {0} ContentSources", lobjContentSources.Count)

        Return lobjContentSources

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      Finally
        'LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try
    End Function

    Public Sub RebuildProviderCatalog()
      Try
        mobjProviderCatalog = GetProviderCatalog()
        Save()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Sub Refresh()
      Try
        GetCurrentSettings(True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

#Region "Singleton Support"

    Public Shared ReadOnly Property Instance As ConnectionSettings
      Get
        Try
          If Not Helper.CallStackContainsMethodName("AssignObjectProperties") Then
            Return GetInstance(False)
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property


    Public Shared Function GetInstance(ByVal lpForceRefresh As Boolean) As ConnectionSettings
      Try
        If lpForceRefresh = True Then
          mobjInstance = GetCurrentSettings(lpForceRefresh)
        ElseIf mobjInstance Is Nothing OrElse mobjInstance.IsDisposed = True Then
          mobjInstance = GetCurrentSettings(lpForceRefresh)
        End If
        'mintReferenceCount += 1
        Return mobjInstance
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "CSF Methods"

    ''' <summary>
    ''' Used to create settings.csf file
    ''' </summary>
    ''' <param name="lpConfigPath">The fully qualified path to the csf file.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function Create(ByVal lpConfigPath As String) As ConnectionSettings
      Try
        ' Create a connection settings object.
        Dim lobjConnectionSettings As New Configuration.ConnectionSettings()

        ' Add all available provider paths
        UpdateCSFProviderPaths(lobjConnectionSettings)

        ' Create the first content source connection string
        Dim lstrDefaultConnectionString As String =
              String.Format(
                "Name=Local File System;Provider=File System Provider;ExportPath={0}Exports\Local File System;ImportPath={0}Exports\Local File System;RootPath=Desktop;UserName=;Password=F8261CB94C60527B;MaxLongFileNameLength=100",
                FileHelper.Instance.CtsDocsPath)

        ' Add the first content source connection string
        lobjConnectionSettings.ContentSourceConnectionStrings.Add(lstrDefaultConnectionString)

        ' Make an entry in the log showing that we created the connection settings file.
        ''ApplicationLogging.WriteLogEntry(String.Format("Created initial connection string '{0}'.",
        ''                                               lstrDefaultConnectionString), TraceEventType.Information, 63482)

        lobjConnectionSettings.mstrSettingsPath = lpConfigPath

        ' Save the file
        SaveCSF(lpConfigPath, lobjConnectionSettings)

        ' Make an entry in the log showing that we created the connection settings file.
        ApplicationLogging.WriteLogEntry(String.Format("Created connection settings file at '{0}'.",
                                                       lpConfigPath), TraceEventType.Information, 63483)

        Return lobjConnectionSettings

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Saves the specified csf file to the specified path.
    ''' </summary>
    ''' <param name="lpSettingsPath">The fully qualified path to save the file to.</param>
    ''' <param name="lpConnectionSettings">The connection settings object reference to work with.</param>
    ''' <remarks></remarks>
    Private Shared Sub SaveCSF(ByVal lpSettingsPath As String,
                               ByVal lpConnectionSettings As Configuration.ConnectionSettings)
      Try
        lpConnectionSettings.Serialize(lpSettingsPath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds all valid and available provider paths to the csf file.
    ''' </summary>
    ''' <param name="lpConfigPath">The fuly qualified path to the csf file.</param>
    ''' <param name="lpConnectionSettings">The connection settings object reference to work with.</param>
    ''' <remarks></remarks>
    Private Shared Sub UpdateCSFProviderPaths(ByVal lpConnectionSettings As Configuration.ConnectionSettings)
      Try
        ' Get all the installed provider files and add them to the provider paths collection
        Dim lstrProviderPaths As String() = IO.Directory.GetFiles(FileHelper.Instance.ProviderPath,
                                                                  "Ecmg.Cts.Providers.*.dll")

        lpConnectionSettings.ProviderPaths.Clear()

        ' Make sure each file is a valid provider file
        For Each lstrProviderPath As String In lstrProviderPaths
          If ContentSource.IsValidProvider(lstrProviderPath) Then
            lpConnectionSettings.ProviderPaths.Add(lstrProviderPath)
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    Private Function GetContentSourceNames() As List(Of String)
      Try
        Dim lobjContentSourceNames As New List(Of String)
        Dim lstrContentSourceName As String '= Helper.GetInfoFromString(mstrConnectionString, "Name")

        For Each lobjContentSourceConnectionString As String In ContentSourceConnectionStrings
          lstrContentSourceName = Helper.GetInfoFromString(lobjContentSourceConnectionString, "Name")
          If Not String.IsNullOrEmpty(lstrContentSourceName) Then
            lobjContentSourceNames.Add(lstrContentSourceName)
          End If
        Next

        Return lobjContentSourceNames

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetProviderCatalog() As Providers.ProviderCatalog
      Try
        'LogSession.EnterMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
        'LogSession.LogDebug("Starting ConnectionSettings.GetProviderCatalog()")
        Dim lobjProviderCatalog As New Providers.ProviderCatalog
        For Each lstrProviderPath As String In ProviderPaths

#If NET8_0_OR_GREATER Then
          If Not lstrProviderPath.Contains("utilities", StringComparison.CurrentCultureIgnoreCase) Then
            Try
              lobjProviderCatalog.Add(lstrProviderPath)
            Catch ex As Exception
              ' We were unable to add this content source
              ' Make a note in the log and move on
              ApplicationLogging.WriteLogEntry _
                (String.Format("Unable to load provider path in GetProviderCatalog '{0}'.",
                               lstrProviderPath), TraceEventType.Error)
            End Try
          End If
#Else
          If Not lstrProviderPath.ToLower.Contains("utilities") Then
            Try
              lobjProviderCatalog.Add(lstrProviderPath)
            Catch ex As Exception
              ' We were unable to add this content source
              ' Make a note in the log and move on
              ApplicationLogging.WriteLogEntry _
                (String.Format("Unable to load provider path in GetProviderCatalog '{0}'.",
                               lstrProviderPath), TraceEventType.Error)
            End Try
          End If
#End If



        Next

        'LogSession.LogDebug("Completed ConnectionSettings.GetProviderCatalog()")

        Return lobjProviderCatalog

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      Finally
        'LogSession.LeaveMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try
    End Function

    Private Sub InitializeJobRunnerExecutionFilePath()
      Try
        'mstrJobRunnerExecutionFilePath
        mstrJobRunnerExecutionFilePath = String.Format("{0}\Jobrunner", FileHelper.Instance.CtsDocsPath.TrimEnd("\"))

        If Not Directory.Exists(mstrJobRunnerExecutionFilePath) Then
          Directory.CreateDirectory(mstrJobRunnerExecutionFilePath)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Private Sub InitializeLogSession(lpSource As String)
    '  Try
    '    mobjLogSession = SiAuto.Si.AddSession("ConnectionSettings")
    '    mobjLogSession.Color = Drawing.Color.Azure

    '    If String.IsNullOrEmpty(SettingsPath) Then
    '      If Not String.IsNullOrEmpty(lpSource) Then
    '        'LogSession.LogMessage("ConnectionSettings initializing from {0}", lpSource)
    '      End If
    '    Else
    '      'LogSession.LogMessage("ConnectionSettings initializing '{0}'", SettingsPath)
    '    End If


    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Private Sub FinalizeLogSession()
    '  Try
    '    'LogSession.LogMessage("ConnectionSettings closing")
    '    SiAuto.Si.DeleteSession(LogSession)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    ' ''' <summary>
    ' '''     Composes the parts from other assemblies that will get folded into this settings file.
    ' ''' </summary>
    ' ''' <remarks>Uses Managed Extensibility Framework</remarks>
    'Private Sub Compose()
    '  Try
    '    mobjAssemblyCatalog = New AssemblyCatalog(Reflection.Assembly.GetExecutingAssembly())
    '    mobjAggregateCatalog = New AggregateCatalog(mobjAssemblyCatalog, New DirectoryCatalog(IO.Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.Location)))

    '    mobjCompositionContainer = New CompositionContainer(mobjAggregateCatalog)
    '    mobjCompositionContainer.SatisfyImportsOnce(Me)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Private Sub MobjObservableContentSourceConnectionStrings_CollectionChanged(sender As Object,
                                                                               e As _
                                                                                Specialized.
                                                                                NotifyCollectionChangedEventArgs) _
      Handles MobjObservableContentSourceConnectionStrings.CollectionChanged
      Try
        ' Syncronize the lists
        'mobjContentSourceConnectionStrings.Clear()
        'mobjContentSourceConnectionStrings.AddRange(mobjObservableContentSourceConnectionStrings.ToArray)
        mobjContentSourceNames = GetContentSourceNames()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ISerialize Implementation"

    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization 
    ''' and deserialization.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DefaultFileExtension() As String Implements ISerialize.DefaultFileExtension
      Get
        Return CONNECTION_SETTINGS_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object _
      Implements ISerialize.Deserialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mstrSettingsPath = lpFilePath

        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex,
                                        String.Format("{0}::Deserialize('{1}', '{2}')", Me.GetType.Name, lpFilePath,
                                                      lpErrorMessage))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function Deserialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.Deserialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Deserialize(lpXML)", Me.GetType.Name))
        Helper.DumpException(ex)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return Serializer.Serialize.Xml(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) _
      Implements ISerialize.Serialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        If lpFileExtension.Length = 0 Then
          ' No override was provided
          If lpFilePath.EndsWith(CONNECTION_SETTINGS_FILE_EXTENSION) = False Then
            lpFilePath = lpFilePath.Remove(lpFilePath.Length - 3) & CONNECTION_SETTINGS_FILE_EXTENSION
          End If

        End If
        mstrSettingsPath = lpFilePath

        ' Check to make sure we can write to the file
        If File.Exists(lpFilePath) Then
          Dim lobjFileInfo As New FileInfo(lpFilePath)
          If lobjFileInfo.IsReadOnly Then
            lobjFileInfo.IsReadOnly = False
            ApplicationLogging.WriteLogEntry(String.Format("Updated {0} file from read-only to read-write.", lpFilePath),
                                             Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information, 62097)
          End If
        End If

        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean,
                         ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mstrSettingsPath = lpFilePath
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ValidateXml(ByVal lpXml As String, ByRef lpErrorMessage As String) As Boolean
      Try
        Return XmlValidator.ValidateXmlString(lpXml, Me.GetType, lpErrorMessage)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub UpdateFromXmlString(ByVal lpXml As String)
      Try
        Dim lstrErrorMessage As String = String.Empty

        If XmlValidator.ValidateXmlString(lpXml, Me.GetType, lstrErrorMessage) = False Then
          ' Try to get the full details
          Dim lobjConnectionSettings As ConnectionSettings = Serializer.Deserialize.XmlString(lpXml, Me.GetType)
        End If

        ' If we made it this far then we will update our actual settings

        ' First we will update the file
        Dim lobjSettingsXmlDocument As New XmlDocument
        lobjSettingsXmlDocument.LoadXml(lpXml)
        lobjSettingsXmlDocument.Save(ConnectionSettings.GetCurrentSettingsPath)

        ' Second we will update the object
        mobjInstance = Serializer.Deserialize.XmlString(lpXml, Me.GetType)

      Catch DeserializeEx As Exceptions.DeserializationException
        ApplicationLogging.LogException(DeserializeEx, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return Serializer.Serialize.XmlString(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Sub Save()
      Try

        If SettingsPath Is Nothing OrElse SettingsPath.Length = 0 Then
          Throw New InvalidOperationException("Cannot save connection settings, there is no settings path set.")
        End If

        Serialize(SettingsPath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region " IDisposable Support "

    Private disposedValue As Boolean = False    ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: free other state (managed objects).
        End If

        ' DISPOSETODO: free your own state (unmanaged objects).
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

#End Region
  End Class
End Namespace
