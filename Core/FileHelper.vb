'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"



Imports System.ComponentModel
Imports System.Configuration
Imports System.Environment
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports Documents.Exceptions
Imports Documents.Transformations
Imports Documents.Utilities

#End Region

Namespace Core
  ''' <summary>
  ''' Singleton Helper class containing only static methods used in Cts specific file operations
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FileHelper

#Region "Class Constants"

    'Private Const CTS_HKEY_PATH As String = "Software\Ecmg\Content Transformation Services"
    'Private Const REG_STR_VAL_INSTALLED_PATH As String = "Installed Path"
    'Private Const REG_STR_VAL_DOCS_ROOT As String = "Documents Root Path"
    Private Const CTS_REPOSITORY_FOLDER_NAME As String = "Repositories"
    Private Const CTS_TRANSFORMATION_FOLDER_NAME As String = "Transformations"
    Private Const LOCAL_SYSTEM_USER As String = "nt authority\system"


    Private Const ERROR_SHARING_VIOLATION As Integer = 32
    Private Const ERROR_LOCK_VIOLATION As Integer = 33

#End Region

#Region "Class Variables"

    Private Shared mstrCurrentUser As String = String.Empty
    Private Shared mobjInstance As FileHelper
    Private Shared mintReferenceCount As Integer
    Private Shared ReadOnly mobjPathDictionary As New Dictionary(Of String, String)
    Private mstrRootPath As String
    Private mstrProviderPath As String
    Private mstrTempPath As String
    'Private mstrLicensePath As String
    Private mstrConfigPath As String
    Private mstrResourcesPath As String
    Private mstrRepositoriesPath As String
    Private mstrRepositoriesTempPath As String
    Private mstrTransformationsPath As String
    Private mstrCtsDocsPath As String
    Private mstrCtsProgramDataPath As String

#End Region

#Region "Public Enumerations"

    Public Enum SpecialPathType
      ''' <summary>
      ''' The Cts root installation folder.
      ''' </summary>
      ''' <remarks>
      ''' Defaults to 'C:\Program Files\ECMG\Content Transformation Services'
      ''' Defined in the registry under 'HKEY_LOCAL_MACHINE\Software\Ecmg\Content Transformation Services\Installed Path'
      ''' </remarks>
      CtsRoot = -1
      ''' <summary>
      ''' The Cts temporary working path.
      ''' </summary>
      ''' <remarks>
      ''' May be defined in the registry under 'HKEY_LOCAL_MACHINE\Software\Ecmg\Content Transformation Services\Temporary Path'
      ''' If not defined in the registry the defaults to CtsDocsPath\Temp
      ''' </remarks>
      CtsTemp = 0
      ''' <summary>
      ''' The Cts configuration folder.
      ''' </summary>
      ''' <remarks></remarks>
      CtsConfig = 1
      ''' <summary>
      ''' The Cts Resources folder.
      ''' </summary>
      ''' <remarks></remarks>
      CtsResources = 2
      ''' <summary>
      ''' The Cts Providers folder.
      ''' </summary>
      ''' <remarks>Deprecated as providers are now located in the CtsRoot path by default.</remarks>
      CtsProviders = 3
      ''' <summary>
      ''' The Cts License folder.
      ''' </summary>
      ''' <remarks></remarks>
      CtsLicenses = 4
      ''' <summary>
      ''' The Cts Repository folder where 
      ''' repository profiles are stored.
      ''' </summary>
      ''' <remarks></remarks>
      CtsRepositories = 5
      ''' <summary>
      ''' The Cts folder used for temporary 
      ''' expansion of repository profiles at runtime.
      ''' </summary>
      ''' <remarks></remarks>
      CtsRepositoryTempArea = 6
      ''' <summary>
      ''' The folder where the current 
      ''' application log is written to.
      ''' </summary>
      ''' <remarks></remarks>
      CtsApplicationLog = 7
      ''' <summary>
      ''' The folder where all program data the user
      ''' does not directly interact with will be 
      ''' stored, read from, and written to.
      ''' </summary>
      ''' <remarks></remarks>
      CtsProgramData = 8
    End Enum

#End Region

#If NET48 Then

#Region "Dll Imports"

    <DllImport("kernel32.dll")>
    Shared Function GetCompressedFileSizeW(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal lpFileName As String,
                                           <Out(), MarshalAs(UnmanagedType.U4)> ByRef lpFileSizeHigh As UInteger) _
      As UInteger
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, PreserveSig:=True)>
    Shared Function GetDiskFreeSpaceW(<[In], MarshalAs(UnmanagedType.LPWStr)> ByVal lpRootPathName As String,
                                      ByRef lpSectorsPerCluster As UInteger, ByRef lpBytesPerSector As UInteger,
                                      ByRef lpNumberOfFreeClusters As UInteger,
                                      ByRef lpTotalNumberOfClusters As UInteger) As Integer
    End Function


#End Region

#End If

#Region "Constructors"

    Private Sub New()
      mintReferenceCount = 0
    End Sub

#End Region

#If NET48 Then

#Region "WINAPI Declarations"

    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function GetShortPathName(ByVal longPath As String,
                                            <MarshalAs(UnmanagedType.LPTStr)> ByVal ShortPath As _
                                             System.Text.StringBuilder,
                                            <MarshalAs(Runtime.InteropServices.UnmanagedType.U4)> ByVal bufferSize As _
                                             Integer) As Integer
    End Function

#End Region

#End If

#Region "Public Properties"

    Public ReadOnly Property CtsDocsPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrCtsDocsPath) Then
            mstrCtsDocsPath = GetCtsDocsPath()
          End If
          Return mstrCtsDocsPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property CtsProgramDataPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrCtsProgramDataPath) Then
            mstrCtsProgramDataPath = GetCtsProgramDataPath()
          End If
          Return mstrCtsProgramDataPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Shared ReadOnly Property ReferenceCount As Integer
      Get
        Return mintReferenceCount
      End Get
    End Property

    Public ReadOnly Property RootPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrRootPath) Then
            mstrRootPath = GetRootPath()
          End If
          Return mstrRootPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property ProviderPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrProviderPath) Then
            mstrProviderPath = GetProviderPath()
          End If
          Return mstrProviderPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property TempPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrTempPath) Then
            mstrTempPath = Path.GetTempPath() 'GetTempPath()
          End If
          Return mstrTempPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    'Public ReadOnly Property LicensePath As String
    '  Get
    '    Try
    '      If String.IsNullOrEmpty(mstrLicensePath) Then
    '        mstrLicensePath = GetLicensePath()
    '      End If
    '      Return mstrLicensePath
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      '  Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

    Public ReadOnly Property ConfigPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrConfigPath) Then
            mstrConfigPath = GetConfigPath()
          End If
          Return mstrConfigPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property ResourcesPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrResourcesPath) Then
            mstrResourcesPath = GetResourcesPath()
          End If
          Return mstrResourcesPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property RepositoriesPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrRepositoriesPath) Then
            mstrRepositoriesPath = GetTargetPath(CTS_REPOSITORY_FOLDER_NAME)
          End If
          Return mstrRepositoriesPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property RepositoriesTempPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrRepositoriesTempPath) Then
            mstrRepositoriesTempPath = GetRepositoriesTempPath()
          End If
          Return mstrRepositoriesTempPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property TransformationPath As String
      Get
        Try
          If String.IsNullOrEmpty(mstrTransformationsPath) Then
            mstrTransformationsPath = GetTargetPath(CTS_TRANSFORMATION_FOLDER_NAME)
          End If
          Return mstrTransformationsPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Private Properties"

    ''' <summary>
    ''' Returns the current user.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Shared ReadOnly Property CurrentUser() As String
      Get
        Try
          If mstrCurrentUser.Length = 0 Then
            mstrCurrentUser = Environment.UserName
          End If
          Return mstrCurrentUser
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Private Shared ReadOnly Property PathDictionary() As Dictionary(Of String, String)
      Get
        Return mobjPathDictionary
      End Get
    End Property

#End Region

#Region "Singleton Support"

    Public Shared Function Instance() As FileHelper
      Try
        If mobjInstance Is Nothing Then
          mobjInstance = New FileHelper
        End If
        mintReferenceCount += 1
        Return mobjInstance
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Methods"

    ''' <summary>Builds a unique file name</summary>
    ''' <param name="lpFolderPath">The foder in which the file will be saved.</param>
    ''' <param name="lpFileType">The type of file to be created.</param>
    ''' <param name="lpExportProviderName">The name of the export provider.</param>
    ''' <param name="lpImportProviderName">The name of the import provider.</param>
    ''' <returns>A fully qualified file name to be used for saving an Ecmg.Cts object.</returns>
    ''' <remarks>Used to generate new unique file names.</remarks>
    Public Shared Function BuildFileName(ByVal lpFolderPath As String,
                                         ByVal lpFileType As FileType,
                                         Optional ByVal lpExportProviderName As String = "",
                                         Optional ByVal lpImportProviderName As String = "") As String

      Dim lstrFileName As String

      Try
        ApplicationLogging.WriteLogEntry("Enter Helper::BuildFileName", TraceEventType.Verbose)

        If lpFolderPath.EndsWith("\"c) Then
          lstrFileName = lpFolderPath
        Else
          lstrFileName = lpFolderPath & "\"
        End If
        'lstrFileName &= lpExportProviderName & "_to_" & lpImportProviderName & Now.ToFileTime & ".clg"
        'lstrFileName &= "_to_" & lpImportProviderName

        Select Case lpFileType
          Case FileType.MigrationLog
            lstrFileName &= lpExportProviderName & "_to_" & lpImportProviderName & Now.Ticks & ".clg"
          Case FileType.MigrationConfigurationFile
            '       lstrLogFile &= "_migration_Con_"
            lstrFileName &= lpExportProviderName & "_to_" & lpImportProviderName & Now.Ticks & ".mcf"
          Case FileType.CondensedMigrationConfiguration
            lstrFileName &= lpExportProviderName & "_to_" & lpImportProviderName & Now.Ticks & ".cmc"
          Case FileType.ValidationResultSetOutputFile
            lstrFileName &= "ValidationResults_" & Now.Ticks & ".vrs"
          Case FileType.ContentTransformationFile
            lstrFileName &= "ContentTransformation_" & Now.Ticks & ".ctf"
          Case FileType.ContentDefinitionFile
            lstrFileName &= "ContentDefinition_" & Now.Ticks & ".cdf"
        End Select

        ApplicationLogging.WriteLogEntry("Exit Helper::BuildFileName", TraceEventType.Verbose)

        Return lstrFileName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("Helper::BuildFileName('{0}', '{1}', '{2}', '{3}')",
                                                          lpFolderPath, lpFileType.ToString, lpExportProviderName,
                                                          lpImportProviderName))
        Return ""
      End Try
    End Function

    Public Shared Function GetCtsAppDataPath() As String
      Try

        Dim lstrCtsAppDataPath As String = String.Empty

        lstrCtsAppDataPath = String.Format("{0}\CTS", GetCompanyAppDataPath()).Replace("\\", "\")

        Return lstrCtsAppDataPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Function GetApplicationConfiguration() As System.Configuration.Configuration
    '  Try
    '    Dim lstrAppConfigName As String = IO.Path.GetFileName(Windows.Forms.Application.ExecutablePath) &
    '                                      ".config"
    '    Dim lstrStartUpPath As String = System.Windows.Forms.Application.StartupPath

    '    Return GetApplicationConfiguration(lstrStartUpPath, lstrAppConfigName)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Given an application config filename, return a configuration object
    '''' </summary>
    '''' <param name="lpConfigFileName"></param>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function GetApplicationConfiguration(ByVal lpExePath As String, ByVal lpConfigFileName As String) _
    '  As System.Configuration.Configuration

    '  Try

    '    If (Not lpExePath.EndsWith("\")) Then
    '      lpExePath &= "\"
    '    End If
    '    Dim lstrSpecialAppConfigPath As String = GetAppConfigPath(lpExePath) & lpConfigFileName
    '    If (Not IO.File.Exists(lstrSpecialAppConfigPath)) Then
    '      'Copy over the "template" from the exe path
    '      IO.File.Copy(lpExePath & lpConfigFileName, lstrSpecialAppConfigPath)
    '    End If

    '    Dim configFile As ExeConfigurationFileMap = New ExeConfigurationFileMap()
    '    configFile.ExeConfigFilename = lstrSpecialAppConfigPath

    '    ApplicationLogging.WriteLogEntry("Using application configuration file: " & lstrSpecialAppConfigPath)

    '    Return ConfigurationManager.OpenMappedExeConfiguration(configFile, ConfigurationUserLevel.None)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Retreives the path used for all app.config files
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Shared Function GetAppConfigPath(ByVal lpApplicationStartupPath As String) As String

    '  Dim lstrCtsAppConfigRootDir As String = String.Empty
    '  Dim lstrCtsAppConfigDir As String = String.Empty

    '  Try

    '    If Helper.IsRunningInstalled = False Then
    '      ' This is not a deployed application, try to load the settings from the shared area
    '      Dim assemblyPath As String = My.Application.Info.DirectoryPath
    '      Dim ctsRootIndex As Integer = assemblyPath.IndexOf("\CTS")
    '      Dim parentDirectory As String = assemblyPath.Remove(ctsRootIndex + 4)
    '      Dim sharedDirectory As String = parentDirectory & "\Shared\"
    '      If System.IO.Directory.Exists(sharedDirectory) Then
    '        lstrCtsAppConfigDir = sharedDirectory
    '        ApplicationLogging.WriteLogEntry("Using application configuration path of: " & lstrCtsAppConfigDir)
    '        Return lstrCtsAppConfigDir
    '      End If
    '    End If


    '    If My.Computer.Info.OSVersion >= "6.0" Then
    '      ' This is Vista or newer, access to the installed folder is limited

    '      ' Use the current user application data directory.
    '      lstrCtsAppConfigRootDir = GetCtsAppDataPath()


    '      lstrCtsAppConfigDir = String.Format("{0}\", lstrCtsAppConfigRootDir)
    '      If IO.Directory.Exists(lstrCtsAppConfigDir) = False Then
    '        IO.Directory.CreateDirectory(lstrCtsAppConfigDir)
    '      End If

    '    Else
    '      ' This is XP or older, we can use the application startup folder
    '      lstrCtsAppConfigDir = lpApplicationStartupPath
    '    End If

    '    Return lstrCtsAppConfigDir

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ''' <summary>
    ''' Gets the Cts Documents Root Path.
    ''' </summary>
    ''' <returns>The root Cts Documents Path.</returns>
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
    Public Function GetCtsDocsPath() As String
      Try

        Dim lstrCtsDocsRootDir As String = String.Empty
        Dim lstrCtsDocsDir As String = String.Empty
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
        '    lstrCtsDocsDir = Helper.RemoveExtraBackSlashesFromFilePath(lstrCtsDocumentsRootPath)
        '  End If
        'End If

        lstrCtsDocsDir = ConfigurationManager.AppSettings("CTSDocsPath") ' Configuration.InstallSettings.Instance.DocumentsRootPath

        'If lstrCtsDocsDir.Length = 0 Then
        If String.IsNullOrEmpty(lstrCtsDocsDir) Then

          If Environment.OSVersion.Platform = PlatformID.Win32NT AndAlso Environment.OSVersion.Version.Major >= 6 Then
            ' This is Vista or newer, access to the installed folder is limited

            'Use the current user application data directory
            lstrCtsDocsRootDir = GetCtsAppDataPath()

            'If CurrentUser <> LOCAL_SYSTEM_USER Then
            '  ' If we are running as a regular named user then use their Documents directory
            '  lstrCtsDocsRootDir = My.Computer.FileSystem.SpecialDirectories.MyDocuments

            'Else
            '  ' If we are running as some thing else such as LOCAL SYSTEM such as with 
            '  ' Content Loader Service then use the current user application data directory.
            '  lstrCtsDocsRootDir = GetCtsAppDataPath()
            'End If

            lstrCtsDocsDir = String.Format("{0}\", lstrCtsDocsRootDir)
            If IO.Directory.Exists(lstrCtsDocsDir) = False Then
              IO.Directory.CreateDirectory(lstrCtsDocsDir)
            End If
          Else
            ' This is XP or older, we can use the installed folder
            lstrCtsDocsDir = RootPath
          End If

        End If

        Return lstrCtsDocsDir

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetCtsProgramDataPath() As String
      Try
        Return GetSpecialPath(SpecialPathType.CtsProgramData)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Gets a specific working path for CTS operations
    ''' </summary>
    ''' <param name="lpSpecialPathType">An enumeration of SpecialPathType specifying the type of path requested</param>
    ''' <returns>A fully qualified directory path</returns>
    ''' <remarks></remarks>
    Public Function GetSpecialPath(ByVal lpSpecialPathType As SpecialPathType) As String

      Dim lstrSpecialPath As String = String.Empty

      Try

        Select Case lpSpecialPathType

          Case SpecialPathType.CtsRoot
            If PathDictionary.ContainsKey(SpecialPathType.CtsRoot) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsRoot)
            Else
              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(RootPath())
              PathDictionary.Add(SpecialPathType.CtsRoot, lstrSpecialPath)
            End If
            'lstrSpecialPath = GetRootPath()

          Case SpecialPathType.CtsTemp
            If PathDictionary.ContainsKey(SpecialPathType.CtsTemp) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsTemp)
            Else
              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(TempPath())
              PathDictionary.Add(SpecialPathType.CtsTemp, lstrSpecialPath)
            End If

          'Case SpecialPathType.CtsLicenses
          '  If PathDictionary.ContainsKey(SpecialPathType.CtsLicenses) Then
          '    lstrSpecialPath = PathDictionary(SpecialPathType.CtsLicenses)
          '  Else
          '    lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(LicensePath())
          '    PathDictionary.Add(SpecialPathType.CtsLicenses, lstrSpecialPath)
          '  End If

          Case SpecialPathType.CtsConfig
            If PathDictionary.ContainsKey(SpecialPathType.CtsConfig) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsConfig)
            Else
              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(ConfigPath())
              PathDictionary.Add(SpecialPathType.CtsConfig, lstrSpecialPath)
            End If

          Case SpecialPathType.CtsProviders
            If PathDictionary.ContainsKey(SpecialPathType.CtsProviders) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsProviders)
            Else
              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(ProviderPath())
              PathDictionary.Add(SpecialPathType.CtsProviders, lstrSpecialPath)
            End If

          Case SpecialPathType.CtsResources
            If PathDictionary.ContainsKey(SpecialPathType.CtsResources) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsResources)
            Else
              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(ResourcesPath())
              PathDictionary.Add(SpecialPathType.CtsResources, lstrSpecialPath)
            End If

          Case SpecialPathType.CtsRepositories
            If PathDictionary.ContainsKey(SpecialPathType.CtsRepositories) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsRepositories)
            Else
              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(RepositoriesPath())
              PathDictionary.Add(SpecialPathType.CtsRepositories, lstrSpecialPath)
            End If

          Case SpecialPathType.CtsRepositoryTempArea
            If PathDictionary.ContainsKey(SpecialPathType.CtsRepositoryTempArea) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsRepositoryTempArea)
            Else
              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(RepositoriesTempPath())
              PathDictionary.Add(SpecialPathType.CtsRepositoryTempArea, lstrSpecialPath)
            End If

          Case SpecialPathType.CtsApplicationLog
            lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(ApplicationLogging.LogDirectory)

          Case SpecialPathType.CtsProgramData
            If PathDictionary.ContainsKey(SpecialPathType.CtsProgramData) Then
              lstrSpecialPath = PathDictionary(SpecialPathType.CtsProgramData)
            Else

              Dim lstrCtsProgramDataPath As String = String.Format("{0}\ECMG\Content Transformation Services",
                                                                   Environment.GetFolderPath(
                                                                     SpecialFolder.CommonApplicationData))

              lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(lstrCtsProgramDataPath)
              If Directory.Exists(lstrSpecialPath) = False Then
                Directory.CreateDirectory(lstrSpecialPath)
              End If

              PathDictionary.Add(SpecialPathType.CtsProgramData, lstrSpecialPath)
            End If
            'Case SpecialPathType.CtsAppConfig
            '  If PathDictionary.ContainsKey(SpecialPathType.CtsAppConfig) Then
            '    lstrSpecialPath = PathDictionary(SpecialPathType.CtsAppConfig)
            '  Else
            '    lstrSpecialPath = Helper.RemoveExtraBackSlashesFromFilePath(GetAppConfigPath())
            '    PathDictionary.Add(SpecialPathType.CtsAppConfig, lstrSpecialPath)
            '  End If

          Case Else
            Throw New ArgumentException(
              String.Format("Unable to get special path: '{0}' is not a valid value for lpSpecialPathType",
                            lpSpecialPathType))

        End Select

        Return lstrSpecialPath

      Catch ArgumentEx As ArgumentException
        If ArgumentEx.Message = "An item with the same key has already been added." Then
          Return lstrSpecialPath
        Else
          ApplicationLogging.LogException(ArgumentEx,
                                          String.Format("FileHelper::GetSpecialPath('{0}')", lpSpecialPathType.ToString))
          ' Re-throw the exception to the caller
          Throw
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex,
                                        String.Format("FileHelper::GetSpecialPath('{0}')", lpSpecialPathType.ToString))
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    '''' <summary>
    '''' Test a directory for create file access permissions
    '''' </summary>
    '''' <param name="lpDirectoryPath">Full path to directory </param>
    '''' <param name="lpAccessRight">File System right tested</param>
    '''' <returns>State</returns>
    'Public Shared Function DirectoryHasPermission(lpDirectoryPath As String, lpAccessRight As FileSystemRights) _
    '  As Boolean
    '  If String.IsNullOrEmpty(lpDirectoryPath) Then
    '    Return False
    '  End If

    '  Try
    '    Dim lobjRules As AuthorizationRuleCollection = Directory.GetAccessControl(lpDirectoryPath).GetAccessRules(True,
    '                                                                                                              True,
    '                                                                                                              GetType(
    '                                                                                                               System _
    '                                                                                                               .
    '                                                                                                               Security _
    '                                                                                                               .
    '                                                                                                               Principal _
    '                                                                                                               .
    '                                                                                                               SecurityIdentifier))
    '    Dim lobjIdentity As WindowsIdentity = WindowsIdentity.GetCurrent()

    '    For Each lobjRule As FileSystemAccessRule In lobjRules
    '      If lobjIdentity.Groups.Contains(lobjRule.IdentityReference) Then
    '        If (lpAccessRight And lobjRule.FileSystemRights) = lpAccessRight Then
    '          If lobjRule.AccessControlType = AccessControlType.Allow Then
    '            Return True
    '          End If
    '        End If
    '      End If
    '    Next
    '  Catch
    '  End Try
    '  Return False
    'End Function

    ''' <summary>
    ''' Takes a folder path and file name and returns the shorter name version of it.
    ''' </summary>
    ''' <param name="lpFolderPath">The candidate folder path.</param>
    ''' <param name="lpFileName">The candidate file name.</param>
    ''' <returns>A valid fully qualified path that is no longer than 255 characters.</returns>
    ''' <remarks>
    ''' This is for use when working with potentially very large file 
    ''' names to ensure that the final path is 255 characters or less.
    ''' </remarks>
    ''' <exception cref="InvalidPathException">
    ''' If the specified folder path is too long to allow a shortened 
    ''' fully qualified path, an InvalidPathException will be thrown.
    ''' </exception>
    Public Shared Function GetValidLengthFilePath(lpFolderPath As String, lpFileName As String) As String
      Try

        ' Make sure we have valid parameter values
        If String.IsNullOrEmpty(lpFolderPath) Then
          Throw New ArgumentNullException(NameOf(lpFolderPath))
        End If

        If String.IsNullOrEmpty(lpFileName) Then
          Throw New ArgumentNullException(NameOf(lpFileName))
        End If


        Dim lintFolderPathLength As Integer = lpFolderPath.Length

        If lintFolderPathLength > 255 Then
          Throw New InvalidPathException("The specified folder path is too long.", lpFolderPath)
        End If

        Dim lintFileNameLength As Integer = lpFileName.Length
        Dim lintTotalCandidatePathLength As Integer = lintFolderPathLength + lintFileNameLength

        If lintTotalCandidatePathLength > 255 Then
          ' We need a shorter file name
          Dim lintMaxAvailableFileLength As Integer = 255 - lintFolderPathLength
          Dim lintExtensionLength As Integer = lintFileNameLength - Path.GetFileNameWithoutExtension(lpFileName).Length -
                                               1

          If lintMaxAvailableFileLength < lintExtensionLength + 2 Then
            Throw _
              New InvalidPathException("The specified folder path is too long to accomodate a shortened file.",
                                       lpFolderPath)
          End If

          Dim lintTrimLength As Integer = lintFileNameLength - lintMaxAvailableFileLength

          Return Path.Combine(lpFolderPath, lpFileName.Substring(lintTrimLength)).Trim()
        Else
          Return Path.Combine(lpFolderPath, lpFileName)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#If NET48 Then

    ''' <summary>
    ''' Takes a long file name and returns the short name version of it.
    ''' </summary>
    ''' <param name="lpFileName">The path to convert to its short representation path.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ShortPathName(ByVal lpFileName As String) As String

      Dim longPathLength As Int32
      Dim lobjShortPathBuilder As StringBuilder
      Dim llngreturnValue As Int32
      Dim lstrShortPathName As String

      Try
        longPathLength = lpFileName.Length + 1
        lobjShortPathBuilder = New StringBuilder(longPathLength)

        ' Call the WINAPI function to do the conversion...
        llngreturnValue = GetShortPathName(lpFileName, lobjShortPathBuilder, longPathLength)

        If llngreturnValue <> 0 Then
          lstrShortPathName = lobjShortPathBuilder.ToString()
        Else
          ' If the return code is zero then there was a problem
          Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        Return lstrShortPathName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End If

    ''' <summary>
    ''' Takes a fully qualified file path and returns only the file name.
    ''' </summary>
    ''' <param name="lpFilePath">Fully qualified file path</param>
    ''' <returns></returns>
    ''' <remarks>Parses based on the '\' folder delimiter</remarks>
    Public Shared Function GetFileName(ByVal lpFilePath As String) As String
      Try
        If lpFilePath.Contains("\"c) = False Then
          Throw _
            New InvalidOperationException(
              String.Format(
                "Unable to get file name from fully qualified path.  The path '{0}' does not contain any '\' delimiters.",
                lpFilePath))
        End If
        Return lpFilePath.Substring(lpFilePath.LastIndexOf("\"c) + 1)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Takes a fully qualified file path and returns only the directory name.
    ''' </summary>
    ''' <param name="lpFilePath">Fully qualified file path</param>
    ''' <returns></returns>
    ''' <remarks>Parses based on the '\' folder delimiter</remarks>
    Public Shared Function GetDirectoryName(ByVal lpFilePath As String) As String
      Try
        If lpFilePath.Contains("\"c) = False Then
          Throw _
            New InvalidOperationException(
              String.Format(
                "Unable to get directory name from fully qualified path.  The path '{0}' does not contain any '\' delimiters.",
                lpFilePath))
        End If
        Return String.Format("{0}\", lpFilePath.Substring(0, lpFilePath.LastIndexOf("\"c)))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function WriteFileToMemoryStream(ByVal lpFilePath As String) As IO.MemoryStream
      Try
        Return Helper.WriteFileToMemoryStream(lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    '''' <summary>
    '''' Goes through all of the providers in the current Providers 
    '''' folder and attempts to add them to the current settings file
    '''' </summary>
    '''' <remarks></remarks>
    'Public Shared Sub AddAllProvidersToSettings()
    '  Try
    '    Configuration.ConnectionSettings.AddAllInstalledProvidersToCurrentSettings()
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Public Shared Sub ClearTempFiles(ByVal lpTargetDirectory As String,
                                     ByVal ParamArray lpTargetExtensions() As String)

      Try

        Dim lstrExtensionCandidate As String = String.Empty

        For Each lstrTargetExtension As String In lpTargetExtensions

          ' If we were only provided 'cdf', change it to '*.cdf'
          If lstrTargetExtension.Contains("."c) = False Then
            lstrTargetExtension = String.Format("*.{0}", lstrTargetExtension)
          End If

          ' If we were only provided '.cdf', change it to '*.cdf'
          If lstrTargetExtension.StartsWith("."c) Then
            lstrTargetExtension = String.Format("*{0}", lstrTargetExtension)
          End If

          For Each lstrFile As String In System.IO.Directory.GetFiles(lpTargetDirectory, lstrTargetExtension)
            Try
              If Path.GetExtension(lstrFile).Equals(".cdf", StringComparison.CurrentCultureIgnoreCase) Then
                ' This is a cdf file, check to see if it is in the temp path and if so, 
                ' clear the contents and content folders as well.
                Dim lobjDocumentToClean As New Document(lstrFile)
                Dim lstrVersionDirectory As String
                Dim lstrVersionFiles As String()
                Dim lstrFileToRemovePath As String
                Dim lstrRelationshipsDirectory As String = String.Format("{0}Relationships\", lpTargetDirectory)

                For Each lobjRelationship As Relationship In lobjDocumentToClean.Relationships
                  If lobjRelationship.RelatedDocument IsNot Nothing Then
                    lstrFileToRemovePath = String.Empty
                    If Not String.IsNullOrEmpty(lobjRelationship.RelatedDocument.SerializationPath) Then
                      lstrFileToRemovePath = lobjRelationship.RelatedDocument.SerializationPath
                    Else
                      lstrFileToRemovePath = String.Format("{0}{1}.cpf", lstrRelationshipsDirectory,
                                                           lobjRelationship.RelatedDocument.ID)
                    End If
                    If File.Exists(lstrFileToRemovePath) Then
                      File.Delete(lstrFileToRemovePath)
                    End If
                  End If
                Next

                ' Try to remove the Relationships directory if it is empty
                If Directory.Exists(lstrRelationshipsDirectory) Then
                  lstrVersionFiles = Directory.GetFiles(lstrRelationshipsDirectory, "*.*")
                  If lstrVersionFiles.Length = 0 Then
                    Directory.Delete(lstrRelationshipsDirectory)
                  End If
                End If

                For Each lobjVersion As Version In lobjDocumentToClean.Versions
                  If lobjVersion.HasContent Then
                    For Each lobjContent As Content In lobjVersion.Contents
                      If File.Exists(lobjContent.CurrentPath) AndAlso
                         lobjContent.CurrentPath.Contains(FileHelper.Instance.TempPath) Then
                        File.Delete(lobjContent.CurrentPath)
                      End If
                    Next
                    ' Try to remove the version directory if it is empty
                    lstrVersionDirectory = String.Format("{0}\{1}", Path.GetDirectoryName(lstrFile), lobjVersion.ID)
                    If Directory.Exists(lstrVersionDirectory) Then
                      lstrVersionFiles = Directory.GetFiles(lstrVersionDirectory, "*.*")
                      If lstrVersionFiles.Length = 0 Then
                        Directory.Delete(lstrVersionDirectory)
                      End If
                    End If
                  End If
                Next
              End If

              ' Delete the file
              System.IO.File.Delete(lstrFile)

            Catch IoEx As IOException
              ApplicationLogging.WriteLogEntry(String.Format("Unable to delete temporary file '{0}'",
                                                             IoEx.Message),
                                               Reflection.MethodBase.GetCurrentMethod,
                                               TraceEventType.Warning, 12342,
                                               lpTargetExtensions)
              ' Just skip it 

            Catch UnauthorizedEx As UnauthorizedAccessException
              ApplicationLogging.WriteLogEntry(String.Format("Unable to delete temporary file '{0}'",
                                                             UnauthorizedEx.Message),
                                               Reflection.MethodBase.GetCurrentMethod,
                                               TraceEventType.Warning, 2,
                                               lpTargetExtensions)
              ' Just skip it 

            Catch ex As Exception
              ApplicationLogging.WriteLogEntry(String.Format("Unable to delete temporary file '{0}'",
                                                             ex.Message),
                                               Reflection.MethodBase.GetCurrentMethod,
                                               TraceEventType.Warning, 777, lpTargetExtensions)
              ' Just skip it

            End Try
          Next
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Just log it and move on
      End Try
    End Sub

    ''' <summary>
    ''' Locates all .cpf files in the specified source 
    ''' folder and extracts the content to the destination folder.
    ''' </summary>
    ''' <param name="lpSourceFolder">The folder path containing the .cpf files.</param>
    ''' <param name="lpDestinationFolder">The folder to which the content should be extracted. 
    ''' If the folder does not exist, we will attempt to create it.</param>
    ''' <param name="lpVersionScope">Specifies which versions of the document to extract the content from.</param>
    ''' <param name="lpOverwriteExistingFiles">
    ''' Specifies whether or not content files already in the 
    ''' destination should be overwritten.  If not, the 
    ''' existing files will remain.</param>
    ''' <remarks></remarks>
    Public Shared Sub ExtractAllContent(ByVal lpSourceFolder As String,
                                        ByVal lpDestinationFolder As String,
                                        ByVal lpVersionScope As VersionScope,
                                        ByVal lpOverwriteExistingFiles As Boolean)
      Try

        ' Make sure the source folder points to a valid folder
        If Directory.Exists(lpSourceFolder) = False Then
          Throw _
            New DirectoryNotFoundException(
              String.Format("The source folder '{0}' specified does not exist or could not be found.", lpSourceFolder))
        End If

        If Directory.Exists(lpDestinationFolder) = False Then
          Try
            ' Attempt to create it
            Directory.CreateDirectory(lpDestinationFolder)
          Catch ex As Exception
            Throw _
              New DirectoryNotFoundException(
                String.Format("The destination folder '{0}' specified does not exist and could not be created.",
                              lpSourceFolder), ex)
          End Try
        End If

        Dim lobjDocument As Document = Nothing
        Dim lstrFilePath As String = String.Empty
        Dim lobjTargetVersion As Version = Nothing

        For Each lstrDocumentFile As String In Directory.GetFiles(lpSourceFolder, "*.cpf")
          ' Make sure we start with an empty reference
          lobjDocument = Nothing
          lobjDocument = New Document(lstrDocumentFile)
          If lobjDocument IsNot Nothing AndAlso lobjDocument.HasContent Then
            Select Case lpVersionScope
              Case VersionScope.AllVersions
                For Each lobjVersion As Version In lobjDocument.Versions
                  If lobjVersion.HasContent Then
                    For Each lobjContent As Content In lobjVersion.Contents
                      lstrFilePath = Helper.CleanPath(String.Format("{0}\{1}", lpDestinationFolder, lobjContent.FileName))
                      lobjContent.WriteToFile(lstrFilePath, lpOverwriteExistingFiles)
                    Next
                  End If
                Next
              Case VersionScope.FirstVersionOnly
                lobjTargetVersion = lobjDocument.FirstVersion
              Case VersionScope.LastVersionOnly
                lobjTargetVersion = lobjDocument.LatestVersion
            End Select

            If lobjTargetVersion IsNot Nothing Then
              ' This should only be the case for first or last version scope
              If lobjTargetVersion.HasContent Then
                For Each lobjContent As Content In lobjTargetVersion.Contents
                  lstrFilePath = Helper.CleanPath(String.Format("{0}\{1}", lpDestinationFolder, lobjContent.FileName))
                  lobjContent.WriteToFile(lstrFilePath, lpOverwriteExistingFiles)
                Next
              End If
            End If
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Shared Function GetLineCount(lpFilePath As String) As Integer
      Try
        Dim lstrLine As String
        Dim lintLineCounter As Integer
        Using lobjSourceFileStream As FileStream = File.Open(lpFilePath, FileMode.Open, FileAccess.Read)
          Using lobjBufferedSourceStream As New BufferedStream(lobjSourceFileStream)
            Using lobjSourceStreamReader As New StreamReader(lobjBufferedSourceStream)
              lstrLine = lobjSourceStreamReader.ReadLine()
              While Not String.IsNullOrEmpty(lstrLine)
                lintLineCounter += 1
                lstrLine = lobjSourceStreamReader.ReadLine()
              End While
            End Using
          End Using
        End Using

        Return lintLineCounter

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Function SplitFile(lpFilePath As String, lpMaxLines As Integer) As List(Of String)
    '  Try

    '    Dim lstrOutputFolder As String = Path.GetDirectoryName(lpFilePath)
    '    If Not FileHelper.DirectoryHasPermission(lstrOutputFolder, FileSystemRights.CreateFiles) Then
    '      lstrOutputFolder = FileHelper.Instance.TempPath
    '    End If

    '    Return FileHelper.SplitFile(lpFilePath, lpMaxLines, lstrOutputFolder)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Public Shared Function SplitFile(lpFilePath As String, lpMaxLines As Integer, lpOutputFolder As String) _
      As List(Of String)
      Try
        Dim lstrNewFileNameBase As String = String.Format("{0}_Split", Path.GetFileNameWithoutExtension(lpFilePath))
        Return SplitFile(lpFilePath, lpMaxLines, lpOutputFolder, lstrNewFileNameBase)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function


    ''' <summary>
    ''' Splits a text file based on the lines in the file.
    ''' </summary>
    ''' <param name="lpFilePath">The path of the file to split.</param>
    ''' <param name="lpMaxLines">The maximum number of lines in the files output by this method.</param>
    ''' <param name="lpOutputFolder">The folder to write the output files to.</param>
    ''' <param name="lpNewFileNameBase">The base file name for the output files.</param>
    ''' <remarks>
    ''' This method is used to break down very large text files (usually id files) into
    ''' smaller pieces so they can be spread out into more manageable sized operations.
    ''' </remarks>
    Public Shared Function SplitFile(lpFilePath As String, lpMaxLines As Integer, lpOutputFolder As String,
                                     lpNewFileNameBase As String) As List(Of String)

      Dim lobjCurrentSplitStreamWriter As StreamWriter = Nothing
      Dim lstrCurrentOutputFileName As String = String.Empty
      Dim lobjOutputFileList As New List(Of String)

      Try
        If Not File.Exists(lpFilePath) Then
          Throw New InvalidPathException(lpFilePath)
        End If

        Dim lstrLine As String
        Dim lintLineCounter As Integer
        Dim lintOutputFileCounter As Integer = 1
        lobjCurrentSplitStreamWriter = GetSplitFile(lpOutputFolder, lpNewFileNameBase, lintOutputFileCounter)
        lstrCurrentOutputFileName = DirectCast(lobjCurrentSplitStreamWriter.BaseStream, FileStream).Name

        Console.WriteLine("Created file '{0}'", lstrCurrentOutputFileName)

        Using lobjSourceFileStream As FileStream = File.Open(lpFilePath, FileMode.Open, FileAccess.Read)
          Using lobjBufferedSourceStream As New BufferedStream(lobjSourceFileStream)
            Using lobjSourceStreamReader As New StreamReader(lobjBufferedSourceStream)
              lstrLine = lobjSourceStreamReader.ReadLine()

              While Not String.IsNullOrEmpty(lstrLine)

                lintLineCounter += 1

                ' See if we are on a split
                If Not lintLineCounter Mod lpMaxLines = 0 Then
                  If Not lobjSourceStreamReader.EndOfStream Then
                    lobjCurrentSplitStreamWriter.WriteLine(lstrLine)
                  Else
                    lobjCurrentSplitStreamWriter.Write(lstrLine.Replace(ControlChars.CrLf, String.Empty))
                  End If
                  lstrLine = lobjSourceStreamReader.ReadLine()
                Else
                  ' We are on a split

                  ' Remove any carriage returns and write the last line
                  lobjCurrentSplitStreamWriter.Write(lstrLine.Replace(ControlChars.CrLf, String.Empty))

                  ' Close the current file and start a new one.
                  lobjCurrentSplitStreamWriter.Close()
                  lobjCurrentSplitStreamWriter.Dispose()
                  lobjOutputFileList.Add(lstrCurrentOutputFileName)
                  lintOutputFileCounter += 1

                  lstrLine = lobjSourceStreamReader.ReadLine()

                  ' If we haven't reached the end of the file yet, create a new file.
                  If Not String.IsNullOrEmpty(lstrLine) Then
                    lobjCurrentSplitStreamWriter = GetSplitFile(lpOutputFolder, lpNewFileNameBase, lintOutputFileCounter)
                    lstrCurrentOutputFileName = DirectCast(lobjCurrentSplitStreamWriter.BaseStream, FileStream).Name
                    Console.WriteLine("Created file '{0}'", lstrCurrentOutputFileName)
                  End If
                End If

                'lstrLine = lobjSourceStreamReader.ReadLine()

              End While
            End Using
          End Using
        End Using

        If Not lobjOutputFileList.Contains(lstrCurrentOutputFileName) Then
          lobjOutputFileList.Add(lstrCurrentOutputFileName)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lobjCurrentSplitStreamWriter IsNot Nothing Then
          lobjCurrentSplitStreamWriter.Close()
          lobjCurrentSplitStreamWriter.Dispose()
        End If

      End Try
      Return lobjOutputFileList
    End Function

    Public Shared Function IsFileLocked(ByVal lpFilePath As String) As Boolean
      'Try-Catch so we dont crash the program and can check the exception
      Try
        'The "using" is important because FileStream implements IDisposable and
        '"using" will avoid a heap exhaustion situation when too many handles  
        'are left undisposed.
        Using lobjFileStream As FileStream = File.Open(lpFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
          lobjFileStream?.Close()
        End Using
      Catch ex As IOException
        'THE FUNKY MAGIC - TO SEE IF THIS FILE REALLY IS LOCKED!!!
        If IsFileLockedEx(ex) Then
          ' do something, eg File.Copy or present the user with a MsgBox - I do not recommend Killing the process that is locking the file
          Return True
        End If
      Finally
      End Try
      Return False
    End Function

#If NET48 Then

    Public Shared Function GetFileSizeOnDisk(ByVal lpFilePath As String) As Long
      Try
        Dim lobjFileInfo As New FileInfo(lpFilePath)
        Dim lintEmptyNumber As UInteger = Nothing
        Dim lintSectorsPerCluster As UInteger = Nothing
        Dim lintBytesPerSector As UInteger = Nothing
        Dim lintResult As Integer = GetDiskFreeSpaceW(lobjFileInfo.Directory.Root.FullName, lintSectorsPerCluster,
                                                      lintBytesPerSector, lintEmptyNumber,
                                                      lintEmptyNumber)
        If lintResult = 0 Then
          Throw New Win32Exception()
        End If
        Dim lintClusterSize As UInteger = lintSectorsPerCluster * lintBytesPerSector
        Dim lintFileSizeHigh As UInteger = Nothing
        Dim lintCompressedFileSize As UInteger = GetCompressedFileSizeW(lpFilePath, lintFileSizeHigh)
        Dim llngSize As Long
        llngSize = CLng(lintFileSizeHigh) << 32 Or lintCompressedFileSize
        Return ((llngSize + lintClusterSize - 1) \ lintClusterSize) * lintClusterSize
      Catch ex As Exception
        ApplicationLogging.WriteLogEntry(String.Format("Error getting file size on disk from file '{0}'.", lpFilePath), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Error, 123654)
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function IsZeroByteFile(ByVal lpFilePath As String) As Boolean
      Try
        If GetFileSizeOnDisk(lpFilePath) = 0 Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End If

#End Region

#Region "Private Methods"

    Private Shared Function GetCompanyAppDataPath() As String
      Try
        Return Environment.GetFolderPath(SpecialFolder.ApplicationData)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function IsFileLockedEx(ByVal lpException As Exception) As Boolean
      Try
        Dim lintErrorCode As Integer = Marshal.GetHRForException(lpException) And ((1 << 16) - 1)
        Return lintErrorCode = ERROR_SHARING_VIOLATION OrElse lintErrorCode = ERROR_LOCK_VIOLATION
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetSplitFile(lpOutputFolder As String, lpNewFileNameBase As String, lpFileCounter As Integer) _
      As StreamWriter
      Try
        Dim lstrOutputFilePath As String = String.Format("{0}\{1}_{2}.txt", lpOutputFolder, lpNewFileNameBase,
                                                         lpFileCounter)

        If File.Exists(lstrOutputFilePath) Then
          'Throw New ItemAlreadyExistsException(lstrOutputFilePath)
          File.Delete(lstrOutputFilePath)
        End If

        Dim lobjOutputStreamWriter As StreamWriter = File.CreateText(lstrOutputFilePath)

        Return lobjOutputStreamWriter

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetRootPath() As String
      Try

        ''Dim lstrAppRootPath As String = String.Format("{0}\", My.Application.Info.DirectoryPath)

        Dim lstrAppRootPath As String = String.Format("{0}\", Path.GetDirectoryName(Assembly.GetEntryAssembly.Location))

        'Dim lobjKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(CTS_HKEY_PATH)

        'If Not lobjKey Is Nothing Then
        '  ' Try to get the value of the Installed Path string value
        '  ' If the entry is not present in the registry this will return Nothing
        '  Dim lstrInstalledPath As String = lobjKey.GetValue(REG_STR_VAL_INSTALLED_PATH)
        '  If String.IsNullOrEmpty(lstrInstalledPath) Then
        '    ' We could not find the entry
        '    ' Write a warning into the log
        '    ApplicationLogging.WriteLogEntry("There is no value named 'Installed Path' defined at 'HKEY_LOCAL_MACHINE\Software\Ecmg\Content Transformation Services'.  Please ensure that an entry is present and a valid value is provided.  The default value should be 'C:\Program Files\ECMG\Content Transformation Services'", TraceEventType.Warning, 404)
        '  Else
        '    ' We found the entry, make sure it exists.
        '    If IO.Directory.Exists(lstrInstalledPath) = False Then
        '      Throw New IO.DirectoryNotFoundException("The path specified in 'HKEY_LOCAL_MACHINE\Software\Ecmg\Content Transformation Services\Installed Path' is invalid.  Make sure a valid path is configured")
        '    End If
        '    ' Make sure it is clean.
        '    If (lstrInstalledPath.EndsWith("\")) = False Then
        '      lstrInstalledPath += "\"
        '    End If

        '    lstrAppRootPath = Helper.RemoveExtraBackSlashesFromFilePath(lstrInstalledPath)

        'End If
        'Else
        'ApplicationLogging.WriteLogEntry("There is no key defined at 'HKEY_LOCAL_MACHINE\Software\Ecmg\Content Transformation Services'.  Please ensure that an entry is present and a valid value for 'Installed Path' is provided.  The default value should be 'C:\Program Files\ECMG\Content Transformation Services'", TraceEventType.Warning, 404)
        'End If

        lstrAppRootPath = AppContext.BaseDirectory ' Configuration.InstallSettings.Instance.InstalledPath

        Try
          If Not Directory.Exists(lstrAppRootPath) Then
            'Throw New Exceptions.InvalidPathException(String.Format("The Installed Path defined in the registry '{0}' does not exist or is invalid", _
            '                                                        lstrAppRootPath), lstrAppRootPath)
            Throw _
              New Exceptions.InvalidPathException(
                String.Format("The defined Installed Path defined '{0}' does not exist or is invalid",
                              lstrAppRootPath), lstrAppRootPath)
          End If

          Return lstrAppRootPath

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 37652)
          Throw New InvalidOperationException("Unable to get Content Transformation Services Installed Path", ex)
        End Try


      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 37653)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Private Function GetTempPath() As String
    '  Try

    '    Dim lstrTemporaryPath As String
    '    Dim lobjKey As Microsoft.Win32.RegistryKey = My.Computer.Registry.LocalMachine.OpenSubKey(CTS_HKEY_PATH)
    '    If Not lobjKey Is Nothing Then
    '      ' Try to get the value of the Temporary Path string value
    '      ' If the entry is not present in the registry this will return Nothing
    '      lstrTemporaryPath = lobjKey.GetValue("Temporary Path")
    '      If (lstrTemporaryPath Is Nothing) Then
    '        ' We could not find the entry
    '        ' Depending on the operating system we will 
    '        ' either use the CtsDocs path or the system defined Temp path

    '        If CurrentUser <> LOCAL_SYSTEM_USER Then
    '          ' If we are running as a regular named user then use their Documents directory.
    '          lstrTemporaryPath = String.Format("{0}Temp\", CtsDocsPath())
    '        Else
    '          ' If we are running as some thing else such as LOCAL SYSTEM such as with 
    '          ' Content Loader Service then use the Window temp directory.
    '          lstrTemporaryPath = My.Computer.FileSystem.SpecialDirectories.Temp
    '        End If

    '      End If
    '    Else
    '      ' The root key for CTS could not be found
    '      ' Use the operating system defined temp path
    '      lstrTemporaryPath = My.Computer.FileSystem.SpecialDirectories.Temp
    '      ApplicationLogging.WriteLogEntry(
    '        String.Format("Defaulting temporary path to '{0}'.  The key '{1}' could not be found in the registry.",
    '                      lstrTemporaryPath,
    '                      CTS_HKEY_PATH), TraceEventType.Warning)
    '    End If

    '    lstrTemporaryPath = Helper.RemoveExtraBackSlashesFromFilePath(lstrTemporaryPath)

    '    Try
    '      If Not My.Computer.FileSystem.DirectoryExists(lstrTemporaryPath) Then
    '        Try
    '          IO.Directory.CreateDirectory(lstrTemporaryPath)
    '        Catch ex As Exception
    '          Throw _
    '            New Exceptions.InvalidPathException(
    '              String.Format("The Temporary Path defined '{0}' does not exist or is invalid",
    '                            lstrTemporaryPath), lstrTemporaryPath)
    '        End Try
    '      End If

    '      ' Make sure it is clean.
    '      If (lstrTemporaryPath.EndsWith("\")) = False Then
    '        lstrTemporaryPath += "\"
    '      End If

    '      Return lstrTemporaryPath
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, String.Format("Helper::GetSpecialPath('{0}')_CreateDirectory'{1}'",
    '                                                        SpecialPathType.CtsTemp.ToString, lstrTemporaryPath))
    '      Throw New InvalidOperationException("Unable to get temporary path", ex)
    '    End Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    'Private Function GetLicensePath() As String
    '  Try

    '    Dim lstrLicensePath As String = String.Format("{0}Licenses\", RootPath)

    '    Try
    '      If Not My.Computer.FileSystem.DirectoryExists(lstrLicensePath) Then
    '        Throw _
    '          New Exceptions.InvalidPathException(
    '            String.Format("The License Path defined '{0}' does not exist or is invalid",
    '                          lstrLicensePath), lstrLicensePath)
    '      End If
    '      Return lstrLicensePath

    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod,
    '                                      String.Format("lstrLicensePath: '{0}'", lstrLicensePath))
    '      Throw New InvalidOperationException("Unable to get application working path", ex)
    '    End Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    '  Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Private Function GetConfigPath() As String
      Try

        Dim lstrConfigPath As String = String.Empty

        If Helper.IsRunningInstalled = False Then
          ' This is not a deployed application, try to load the settings from the shared area
          Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetEntryAssembly.Location)  'My.Application.Info.DirectoryPath
          'Dim parentDirectory As String = System.IO.Directory.GetParent(assemblyPath).Parent.FullName
          Dim ctsRootIndex As Integer = assemblyPath.IndexOf("\CTS")
          Dim parentDirectory As String = assemblyPath.Remove(ctsRootIndex + 4)
          Dim sharedDirectory As String = parentDirectory & "\Shared"
          'directoryInfo = System.IO.Directory.GetParent(assemblyPath).Parent.FullName
          If System.IO.Directory.Exists(sharedDirectory) Then
            lstrConfigPath = sharedDirectory
          End If
        End If

        If lstrConfigPath = String.Empty Then
          If CurrentUser <> LOCAL_SYSTEM_USER Then
            ' <Modified by: Ernie at 11/26/2012-8:31:16 AM on machine: ERNIE-THINK>
            '	lstrConfigPath = String.Format("{0}Config\", CtsDocsPath())
            lstrConfigPath = String.Format("{0}\Config\", GetSpecialPath(SpecialPathType.CtsProgramData))
            ' </Modified by: Ernie at 11/26/2012-8:31:16 AM on machine: ERNIE-THINK>
          Else
            lstrConfigPath = String.Format("{0}Ecmg\Config\",
                                           Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
          End If
          'lstrConfigPath = String.Format("{0}Config\", GetSpecialPath(SpecialPathType.CtsRoot))
          'Added by RKS, if we want to get/store the settings.csf file on a per user basis then this will return a path to the current user's application data folder
          'lstrConfigPath = String.Format("{0}\Ecmg", Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
        End If

        Try
          If Not Directory.Exists(lstrConfigPath) Then
            'Throw New Exceptions.InvalidPathException(String.Format("The Configuration Path '{0}' does not exist or is invalid", _
            '                                                        lstrConfigPath), lstrConfigPath)
            IO.Directory.CreateDirectory(lstrConfigPath)
          End If
          Return lstrConfigPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Helper::GetSpecialPath('{0}')_CreateDirectory'{1}'",
                                                            SpecialPathType.CtsConfig.ToString, lstrConfigPath))
          Throw New InvalidOperationException("Unable to set application working path", ex)
        End Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetResourcesPath() As String
      Try

        Try
          'Dim lstrAppResourcePath As String = String.Format("{0}Resources\", GetSpecialPath(SpecialPathType.CtsRoot))
          Dim lstrOriginalAppResourcePath As String = String.Format("{0}Resources\", RootPath)
          Dim lstrAppResourcePath As String = String.Format("{0}Resources\", CtsDocsPath())
          Dim lstrTargetFile As String = String.Empty

          Try
            If Not Directory.Exists(lstrOriginalAppResourcePath) Then
              Throw _
                New Exceptions.InvalidPathException(
                  String.Format("The Resources Path '{0}' does not exist or is invalid",
                                lstrOriginalAppResourcePath), lstrOriginalAppResourcePath)
            End If
            'Return lstrAppResourcePath

            ' We actually want to use the Resources folder under the users Documents Root Path
            ' If it does not exist we will create it
            If IO.Directory.Exists(lstrAppResourcePath) = False Then
              IO.Directory.CreateDirectory(lstrAppResourcePath)
            End If

            ' If the files are not all there we will syncronize them
            For Each lstrFile As String In IO.Directory.GetFiles(lstrOriginalAppResourcePath,
                                                                 "*.*", IO.SearchOption.TopDirectoryOnly)
              lstrTargetFile = String.Format("{0}{1}", lstrAppResourcePath, IO.Path.GetFileName(lstrFile))
              If IO.File.Exists(lstrTargetFile) = False Then
                IO.File.Copy(lstrFile, lstrTargetFile)
              End If
            Next

            ' Return the new path
            Return lstrAppResourcePath

          Catch ex As Exception
            ApplicationLogging.LogException(ex, String.Format("Helper::GetSpecialPath('{0}')_CreateDirectory'{1}'",
                                                              SpecialPathType.CtsResources.ToString, lstrAppResourcePath))
            Throw New InvalidOperationException("Unable to set application working path", ex)
          End Try
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Helper::GetSpecialPath('{0}')_GetResourcePath",
                                                            SpecialPathType.CtsResources.ToString))
          Throw New InvalidOperationException("Unable to get application resource path", ex)
        End Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetSharedPath() As String
      Try

        Dim lstrAssemblyPath As String = Path.GetDirectoryName(Assembly.GetEntryAssembly.Location)
        Dim lstrSharedPath As String = String.Empty

#If NET8_0_OR_GREATER Then
        If lstrAssemblyPath.Contains("\cts\", StringComparison.CurrentCultureIgnoreCase) Then
          lstrSharedPath = lstrAssemblyPath.Remove(lstrAssemblyPath.ToLower.LastIndexOf("\cts\") + 5) & "Shared"
        End If
#Else
        If lstrAssemblyPath.ToLower.Contains("\cts\") Then
          lstrSharedPath = lstrAssemblyPath.Remove(lstrAssemblyPath.ToLower.LastIndexOf("\cts\") + 5) & "Shared"
        End If
#End If

        If IO.Directory.Exists(lstrSharedPath) Then
          Return lstrSharedPath
        Else
          'Return String.Empty
          ' Temporarily changing this to just return the assembly path to better align
          ' with applications that are not deployed via a formal installation but 
          ' rather via copying the files to a computer and running
          Return lstrAssemblyPath
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetProviderPath() As String
      Try

        ' First determine whether or not we are running a deployed instance or not.
        ' This is based on the assumption that if the application path is in a 'bin'
        '   directory that we are not deployed but are in the development environment.
        Dim assemblyPath As String = Path.GetDirectoryName(Assembly.GetEntryAssembly.Location) 'My.Application.Info.DirectoryPath
        Dim lstrProvidersPath As String = ""
        If assemblyPath.Contains("bin") Then

          ' This is not a deployed application, try to load the settings from the shared area
          Dim parentDirectory As String = System.IO.Directory.GetParent(assemblyPath).Parent.FullName
          'Dim sharedDirectory As String = parentDirectory '& "\Shared"
          Dim sharedDirectory As String = GetSharedPath()
          'directoryInfo = System.IO.Directory.GetParent(assemblyPath).Parent.FullName
          If System.IO.Directory.Exists(sharedDirectory) Then
            lstrProvidersPath = sharedDirectory
          End If
        Else
          ' This is a deployed application, we will get the location of the deployed config path.
          lstrProvidersPath = String.Format("{0}Providers\", RootPath)

          lstrProvidersPath = Helper.RemoveExtraBackSlashesFromFilePath(lstrProvidersPath)

          ' Since we flattened the directory structure out and no longer have a dedicated provider folder, 
          ' we will just return the CTS root path
          lstrProvidersPath = RootPath()
        End If
        Try
          If Not Directory.Exists(lstrProvidersPath) Then
            Throw _
              New Exceptions.InvalidPathException(String.Format("The Providers Path '{0}' does not exist or is invalid",
                                                                lstrProvidersPath), lstrProvidersPath)
          End If
          Return lstrProvidersPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Helper::GetSpecialPath('{0}')_CreateDirectory'{1}'",
                                                            SpecialPathType.CtsProviders.ToString, lstrProvidersPath))
          Throw New InvalidOperationException("Unable to set application working path", ex)
        End Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetRepositoriesTempPath() As String
      Try

        Dim lstrRepositoryTempPath As String = String.Format("{0}\Temp", RepositoriesPath)

        If IO.Directory.Exists(lstrRepositoryTempPath) = False Then
          IO.Directory.CreateDirectory(lstrRepositoryTempPath)
        End If

        Return lstrRepositoryTempPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetTargetPath(ByVal lpFolderName As String) As String
      Try
        Dim lstrTargetPath As String = String.Format("{0}{1}", CtsDocsPath(), lpFolderName)

        If IO.Directory.Exists(lstrTargetPath) = False Then
          IO.Directory.CreateDirectory(lstrTargetPath)
        End If

        Return lstrTargetPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region
  End Class
End Namespace
