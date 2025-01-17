'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Data
Imports System.Text.RegularExpressions
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Migrations
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Providers
  ''' <summary>Defines an instance of a Provider.</summary>
  ''' <remarks>This object is analogous to a DSN (data source name) in ODBC.</remarks>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ContentSource
    Implements IComparable
    Implements ISerialize
    Implements IDescription
    Implements IEquatable(Of ContentSource)
    Implements IRepositoryConnection
    Implements ITableable
    Implements IDisposable

#Region "Class Constants"

    Public Const IPROVIDER As String = "IProvider"

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const CONTENT_SOURCE_FILE_EXTENSION As String = "csf"

#End Region

#Region "Class Variables"

    Private mblnActive As Boolean = True
    Private mstrName As String = String.Empty
    Private mstrDescription As String = String.Empty
    Private mstrProviderName As String = String.Empty
    Private mstrSerializationPath As String
    Private mobjProviderProperties As New ProviderProperties
    Private mstrConnectionString As String = String.Empty
    Private mstrExportPath As String = String.Empty
    Private mstrImportPath As String = String.Empty
    Private mstrProviderPath As String = String.Empty
    Private WithEvents MobjProvider As IProvider
    Private menuState As ProviderConnectionState = ProviderConnectionState.Disconnected
    Private mobjSecurityToken As Object = Nothing
    Private mobjRepository As Repository = Nothing

    'Private mobjLogSession As Gurock.SmartInspect.Session = Nothing

    ' For IExplorer
    Private mobjSelectedFolder As IFolder

#End Region

#Region "Public Events"

    Public Event ConnectionStateChanged As ConnectionStateChangedEventHandler

#End Region

#Region "Constructors"

    Public Sub New()
      Try
        'InitializeLogSession()
        Properties = New ProviderProperties
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Initialize a new ContentSource using the Name, Provider Name and the Provider Path
    ''' </summary>
    ''' <param name="lpName"></param>
    ''' <param name="lpProviderName"></param>
    ''' <param name="lpProviderPath"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpName As String, ByVal lpProviderName As String, ByVal lpProviderPath As String)
      Try
        'InitializeLogSession()
        Name = lpName
        ProviderName = lpProviderName
        ProviderPath = lpProviderPath

        'CheckProviderLicense()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

    ''' <summary>
    ''' Initialize a new ContentSource using the Name, Provider Name and a collection of Provider Properties
    ''' </summary>
    ''' <param name="lpName"></param>
    ''' <param name="lpProviderName"></param>
    ''' <param name="lpProperties"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpName As String, ByVal lpProviderName As String, ByVal lpProperties As ProviderProperties)
      Try
        'InitializeLogSession()
        Name = lpName
        ProviderName = lpProviderName
        Properties = lpProperties

        'CheckProviderLicense()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

    ''' <summary>
    ''' Initialize a new ContentSource using the connection string
    ''' </summary>
    ''' <param name="lpConnectionString"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpConnectionString As String)
      Try
        'InitializeLogSession()
        mstrConnectionString = lpConnectionString
        InitializePropertiesFromConnectionString()
        'If Provider.ConnectionString = "" Then
        '  Provider.ConnectionString = lpConnectionString
        'End If

        'CheckProviderLicense()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ContentSource::New('{0}')", lpConnectionString))
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    ''' <summary>
    '''  Initialize a new ContentSource using the connection string and security token
    ''' </summary>
    ''' <param name="lpConnectionString"></param>
    ''' <param name="lpSecurityToken"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpConnectionString As String, ByVal lpSecurityToken As Object)

      Try
        'InitializeLogSession()
        mstrConnectionString = lpConnectionString
        InitializePropertiesFromConnectionString()
        'CheckProviderLicense()
        mobjSecurityToken = lpSecurityToken
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ContentSource::New('{0}')", lpConnectionString))
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    ''' <summary>
    ''' Initialize a new ContentSource using the ContentSourceInfo structure
    ''' </summary>
    ''' <param name="lpContentSourceInfo"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpContentSourceInfo As ContentSourceInfo)
      Try
        'InitializeLogSession()
        mstrConnectionString = lpContentSourceInfo.ConnectionString
        InitializePropertiesFromConnectionString()
        Me.Name = lpContentSourceInfo.Name

        'CheckProviderLicense()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

#End Region

#Region "Public Properties"

    <XmlAttribute()>
    Public Property Active As Boolean Implements IRepositoryConnection.Active
      Get
        Return mblnActive
      End Get
      Set(ByVal value As Boolean)
        mblnActive = value
      End Set
    End Property

    <XmlAttribute()>
    Public Property Name() As String Implements IDescription.Name, INamedItem.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    <XmlAttribute()>
    Public Property Description() As String Implements Core.IDescription.Description
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        mstrDescription = value
      End Set
    End Property

    Public Property ProviderName() As String
      Get
        If (mstrProviderName Is Nothing) OrElse (mstrProviderName.Length = 0) Then
          ' Try to get the value from the Provider Properties
          Try

            If Properties IsNot Nothing Then

              Dim lobjProviderProperty As ProviderProperty = Properties("Provider")

              If lobjProviderProperty IsNot Nothing Then
                mstrProviderName = lobjProviderProperty.PropertyValue
              Else
                mstrProviderName = String.Empty
              End If
            Else
              mstrProviderName = String.Empty
            End If

          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
            ' We failed, just go with what we have
          End Try
        End If

        Return mstrProviderName

      End Get
      Set(ByVal value As String)
        mstrProviderName = value
      End Set
    End Property

    Public Property Properties() As ProviderProperties
      Get
        Return mobjProviderProperties
      End Get
      Set(ByVal value As ProviderProperties)
        mobjProviderProperties = value
      End Set
    End Property

    Public ReadOnly Property ConnectionProperties As Core.IProperties Implements IRepositoryConnection.Properties
      Get
        Return Me.Properties
      End Get
    End Property

    Public Property ConnectionString() As String Implements IRepositoryConnection.ConnectionString
      Get
        Return mstrConnectionString
      End Get
      Set(ByVal value As String)
        mstrConnectionString = value
      End Set
    End Property

    Public ReadOnly Property Repository As Repository Implements IRepositoryConnection.Repository
      Get
        Return mobjRepository
      End Get
    End Property

    Public Property ExportPath() As String Implements IRepositoryConnection.ExportPath
      Get
        Return ProperExportPath(False)
      End Get
      Set(value As String)
        mstrExportPath = value
      End Set
    End Property

    Public Property ProperExportPath(Optional ByVal lpEnsureEndSlash As Boolean = False) As String
      Get
        Try
          If (mstrExportPath Is Nothing) OrElse (mstrExportPath.Length = 0) Then
            ' Try to get the value from the Provider Properties
            Try
              If Properties Is Nothing OrElse Properties.Contains(CProvider.EXPORT_PATH) = False Then
                mstrExportPath = String.Empty
                Return mstrExportPath
              End If

              Dim mobjExportProperty As ProviderProperty = Properties(CProvider.EXPORT_PATH)
              If mobjExportProperty IsNot Nothing Then
                mstrExportPath = mobjExportProperty.PropertyValue
              End If
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              ' We failed, just go with what we have
            End Try
          End If
          If mstrExportPath.Length = 0 Then
            ' Just return what we have
            Return mstrExportPath
          End If
          If lpEnsureEndSlash = True Then
            If Right(mstrExportPath, 1) <> "\" Then
              Return mstrExportPath & "\"
            Else
              Return mstrExportPath
            End If
          Else
            Return mstrExportPath
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return mstrExportPath
        End Try
      End Get
      Set(ByVal value As String)
        mstrExportPath = value
      End Set
    End Property

    Public Property ImportPath() As String
      Get
        Try
          ApplicationLogging.WriteLogEntry("Enter ContentSource::Get_ImportPath", TraceEventType.Verbose)
          If (mstrImportPath Is Nothing) OrElse (mstrImportPath.Length = 0) Then
            ' Try to get the value from the Provider Properties
            Try

              If Properties Is Nothing OrElse Properties.Contains(CProvider.IMPORT_PATH) = False Then
                mstrImportPath = String.Empty
                Return mstrImportPath
              End If

              Dim lobjProviderProperty As ProviderProperty = Properties(CProvider.IMPORT_PATH)

              If lobjProviderProperty IsNot Nothing Then
                mstrImportPath = lobjProviderProperty.PropertyValue
              Else
                mstrImportPath = String.Empty
              End If

            Catch InvCastEx As InvalidCastException
              ' We failed, just go with what we have
              Return ""
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              ' We failed, just go with what we have
            End Try
          End If
          Return mstrImportPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return mstrImportPath
        Finally
          ApplicationLogging.WriteLogEntry("Exit ContentSource::Get_ImportPath", TraceEventType.Verbose)
        End Try
      End Get
      Set(ByVal value As String)
        mstrImportPath = value
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

    Public Property ProviderPath() As String
      Get
        Try
          ' ApplicationLogging.WriteLogEntry("Enter ContentSource::Get_ProviderPath", TraceEventType.Verbose)
          If mstrProviderPath.Length = 0 Then
            Try

              Dim lobjProviderProperty As ProviderProperty = Properties("ProviderPath")

              If lobjProviderProperty IsNot Nothing Then
                mstrProviderPath = lobjProviderProperty.PropertyValue
              Else
                mstrProviderPath = String.Empty
              End If

            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              mstrProviderPath = ""
            End Try
          End If
          Return mstrProviderPath
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return mstrProviderPath
        Finally
          ' ApplicationLogging.WriteLogEntry("Exit ContentSource::Get_ProviderPath", TraceEventType.Verbose)
        End Try
      End Get
      Set(ByVal value As String)
        Try
          If value IsNot Nothing AndAlso value.Length > 0 AndAlso IO.File.Exists(value) = False Then
            'Throw New InvalidPathException("The supplied value is not a valid path.", value)
            ApplicationLogging.WriteLogEntry(String.Format("The provider path '{0}' given for the content source is not a valid path.  The content source will not be able to connect to the repository", value), TraceEventType.Warning, 6404)
          End If
          mstrProviderPath = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property Provider() As IProvider Implements IRepositoryConnection.Provider
      Get
        Try
          If MobjProvider Is Nothing Then
            MobjProvider = GetProvider(False)
          End If
          Return MobjProvider
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          'Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property Provider(ByVal bAutoConnect As Boolean) As IProvider
      Get
        Try
          If MobjProvider Is Nothing Then
            MobjProvider = GetProvider(bAutoConnect)
          End If
          Return MobjProvider
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          'Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property SupportsClassification As Boolean
      Get
        Try
          Return Me.Provider.SupportsInterface(ProviderClass.Classification)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <XmlIgnore()>
    Public Property SelectedFolder() As IFolder
      Get
        Return mobjSelectedFolder
      End Get
      Protected Friend Set(ByVal value As IFolder)
        mobjSelectedFolder = value
        Provider.SelectedFolder = value
      End Set
    End Property

    ''' <summary>
    ''' Used to validate that the referenced content 
    ''' repository is available for use by the content source.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    Public Property State As ProviderConnectionState Implements IRepositoryConnection.State
      Get
        Try
          Return menuState
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Private Set(value As ProviderConnectionState)
        menuState = value
      End Set
    End Property

    Public ReadOnly Property SecurityToken() As Object
      Get
        Try
          Return mobjSecurityToken
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          'Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Public Methods"

#Region "RIF Generation"

    '''' <summary>
    '''' Gets the default Repository path.
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GetDefaultRifPath() As String
    '  Try
    '    'Dim lstrRifDirectory As String = String.Format("{0}Rifs", FileHelper.GetSpecialPath(SpecialPathType.CtsRoot))
    '    'Dim lstrRifDirectory As String = String.Format("{0}Rifs", FileHelper.GetCtsDocsPath())
    '    Dim lstrRifDirectory As String = FileHelper.Instance.RepositoriesPath
    '    'If IO.Directory.Exists(lstrRifDirectory) = False Then
    '    '  IO.Directory.CreateDirectory(lstrRifDirectory)
    '    'End If

    '    Dim lstrFileName As String = String.Format("{0}\{1}.rfa", lstrRifDirectory, Me.Name)

    '    Return lstrFileName

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Gets the default Repository temp path.
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GetDefaultRifTemporaryPath() As String
    '  Try
    '    Dim lstrRifTempDirectory As String = FileHelper.Instance.RepositoriesTempPath

    '    Dim lstrFileName As String = String.Format("{0}\{1}.rfa", lstrRifTempDirectory, Me.Name)

    '    Return lstrFileName

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ''' <summary>
    ''' Generates and returns a Repository object from the current ContentSource
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRepository() As Repository Implements IRepositoryConnection.GetRepository
      Try

        Dim lstrClasses As String = String.Empty
        Dim lstrLibraryProperties As String = String.Empty
        Dim lstrRequiredProperties As String = String.Empty

        If Me.Provider Is Nothing Then
          Throw New Exceptions.ProviderNotInitializedException
        End If

        ' If the provider does not support classification then we can't create a repository object.
        If Me.Provider.SupportsInterface(ProviderClass.Classification) = False Then
          Throw New Exception("This content source does not support classification.")
        End If

        ' Make sure the provider is connected before attempting to create a repository object.
        If Me.Provider.IsConnected = False Then
          Try
            Me.Provider.Connect(Me)
          Catch RepositoryEx As RepositoryNotAvailableException
            Throw
          Catch ex As Exception
            Throw New Exceptions.RepositoryNotAvailableException(Me.Name, ex.Message, ex)
          End Try
        End If

        Dim lobjRepositoryInformation As New Repository(Me.Name, Me, Me.Provider)

        mobjRepository = lobjRepositoryInformation

        Return lobjRepositoryInformation

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Public Function IsAvailable() As Boolean
      Try
        Select Case State
          Case ProviderConnectionState.Connected
            Return True
          Case ProviderConnectionState.Unavailable
            Return False
          Case ProviderConnectionState.Disconnected
            Try
              Provider.Connect(Me)
              Select Case State
                Case ProviderConnectionState.Connected
                  Return True
                Case ProviderConnectionState.Unavailable
                  Return False
                Case ProviderConnectionState.Disconnected
                  Return False
              End Select
            Catch ex As Exception
              Return False
            End Try
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    '''' <summary>
    '''' Generates a RIF file and returns the location of the file
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GenerateRIF() As String
    '  Try

    '    Return GenerateRIF(GetDefaultRifPath.Replace(".rfa", ".rif"))

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    '''' <summary>
    '''' Generates a RIF file and returns the location of the file
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Function GenerateRIF(ByRef lpRepository As Repository) As String
    '  Try

    '    Return GenerateRIF(GetDefaultRifPath.Replace(".rfa", ".rif"), lpRepository)

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    ''' <summary>
    ''' Generates a RIF file and returns the location of the file
    ''' </summary>
    ''' <param name="lpSerializationPath">The fully qualified file name to save the RIF to.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GenerateRIF(ByVal lpSerializationPath As String) As String
      Try

        Dim lobjRepository As Repository = GetRepository()


        ' Use the signature with the extension option so we can get back any updated filename
        ' This is in case the rif gets zipped, we will get back the correct file name
        lobjRepository?.Serialize(lpSerializationPath, True, True)

        Return lpSerializationPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Generates a RIF file and returns the location of the file
    ''' </summary>
    ''' <param name="lpSerializationPath">The fully qualified file name to save the RIF to.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GenerateRIF(ByVal lpSerializationPath As String, ByRef lpRepository As Repository) As String
      Try

        Dim lobjRepository As Repository = GetRepository()


        ' Use the signature with the extension option so we can get back any updated filename
        ' This is in case the rif gets zipped, we will get back the correct file name
        lobjRepository?.Serialize(lpSerializationPath, True, True)

        ' Set the return parameter
        lpRepository = lobjRepository

        Return lpSerializationPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    Public Sub Connect() Implements IRepositoryConnection.Connect
      Try
        If Me.Provider IsNot Nothing AndAlso Me.State = ProviderConnectionState.Disconnected Then
          Me.Provider.Connect(Me)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Disconnect() Implements IRepositoryConnection.Disconnect
      Try
        If Me.Provider IsNot Nothing AndAlso Me.State = ProviderConnectionState.Connected Then
          Me.Provider.Disconnect()
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overloads Function Equals(ByVal other As ContentSource) As Boolean Implements System.IEquatable(Of ContentSource).Equals
      Try
        ' Compare the parts
        ' Name

        If other Is Nothing Then
          Return False
        End If

        If other.Name <> Me.Name Then
          Return False
        End If

        ' Provider Name
        If other.ProviderName <> Me.ProviderName Then
          Return False
        End If

        ' ConnectionString
        If other.ConnectionString <> Me.ConnectionString Then
          Return False
        End If

        '  Begin changed by Ernie Bahr
        '  June 4th, 2008
        '  We do not need to look into the provider to effectively test for equality.
        '  This will force a provider initialization and we want to minimize the number of times we do that.
        '' Make sure we can get a reference to the provider
        'If lobjContentSource.Provider Is Nothing Then
        '  Throw New ProviderNotInitializedException(Me)
        'End If

        '' ExportPath
        'If lobjContentSource.Provider.SupportsInterface(ProviderClass.Exporter) AndAlso lobjContentSource.ExportPath <> Me.ExportPath Then
        '  Return False
        'End If

        '' ImportPath
        'If lobjContentSource.Provider.SupportsInterface(ProviderClass.Importer) AndAlso lobjContentSource.ImportPath <> Me.ImportPath Then
        '  Return False
        'End If

        '  End changed by Ernie Bahr
        '  June 4th, 2008

        ' ProviderProperties
        If other.Properties.Equals(Me.Properties) = False Then
          Return False
        End If

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean

      Dim lobjContentSource As ContentSource

      ' Make sure it is a real ContentSource
      Try
        'lobjContentSource = CType(obj, ContentSource)
        lobjContentSource = TryCast(obj, ContentSource)

        If lobjContentSource Is Nothing Then
          Return False
        Else
          Return Equals(lobjContentSource)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try

    End Function

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
        Return CONTENT_SOURCE_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.DeSerialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Helper.DumpException(ex)
        Return Nothing
      End Try
    End Function

    ''' <summary>
    ''' Saves a representation of the object in an XML file.
    ''' </summary>
    Public Sub Serialize(ByVal lpFilePath As String) Implements ISerialize.Serialize
      Try
        Serialize(lpFilePath, "")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub Serialize(ByRef lpFilePath As String, ByVal lpFileExtension As String) Implements ISerialize.Serialize

      Try
        SetSerializationPath(lpFilePath)
        Serializer.Serialize.XmlFile(Me, lpFilePath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
      'mobjProviderConfiguration = New ProviderConfiguration

      'With mobjProviderConfiguration
      '  .ClassType = Me.ClassType
      '  .ConnectionString = MyBase.ConnectionString
      '  .ExportPath = Me.ExportPath
      '  .ImportPath = Me.ImportPath
      '  .ProviderSystem = Me.ProviderSystem
      'End With

      'Serializer.Serialize.XmlFile(mobjProviderConfiguration, lpFilePath)

    End Sub

    Public Sub Serialize(ByVal lpFilePath As String, ByVal lpWriteProcessingInstruction As Boolean, ByVal lpStyleSheetPath As String) Implements ISerialize.Serialize

      Try
        SetSerializationPath(lpFilePath)
        'Serializer.Serialize.XmlFile(Me, lpFilePath, , mstrXMLProcessingInstructions)
        If lpWriteProcessingInstruction = True Then
          Serializer.Serialize.XmlFile(Me, lpFilePath, , , True, lpStyleSheetPath)
        Else
          Serializer.Serialize.XmlFile(Me, lpFilePath)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return Nothing
      End Try
    End Function

    Public ReadOnly Property SerializationPath() As String
      Get
        Return mstrSerializationPath
      End Get
    End Property

    Public Function SetSerializationPath(ByVal lpSerializationPath As String) As String

      Try
        mstrSerializationPath = lpSerializationPath
        Return SerializationPath
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lpSerializationPath
      End Try

    End Function

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlElementString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return "ContentSource"
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo

      If TypeOf obj Is ContentSource Then
        Return Name.CompareTo(obj.Name)
      Else
        Throw New ArgumentException("Object is not a ContentSource")
      End If

    End Function

#End Region

#Region "ITableable Implementation"

    Public Function ToDataTable() As System.Data.DataTable Implements ITableable.ToDataTable
      Try
        Dim lobjDataTable As New DataTable(String.Format("tbl{0}", Me.GetType.Name))

        lobjDataTable.Columns.Add("PropertyName")
        lobjDataTable.Columns.Add("PropertyValue")

        For Each lobjProviderProperty As ProviderProperty In Me.Properties
          lobjDataTable.Rows.Add(lobjProviderProperty.PropertyName, lobjProviderProperty.PropertyValue)
        Next

        Return lobjDataTable

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Protected Friend Function DebuggerIdentifier() As String
      Try
        Return String.Format("{0} - {1}", ProviderName, Name)
      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function FindProvider(ByVal lpTargetFolder As String,
                                  ByVal lpContentSource As ContentSource) As CProvider

      Dim lstrProviderCandidatePath As String = String.Empty

      Try

        Dim lobjAssembly As System.Reflection.Assembly
        Dim lobjProviderCandidate As Type
        Dim lobjProvider As Providers.CProvider

        For Each lstrProviderCandidatePath In IO.Directory.GetFiles(lpTargetFolder, "Ecmg.Cts.Provider*.dll")
          ' Test each file to see if it is a provider
          ' If it is, then see if it is the right provider.
          Try
            'lobjAssembly = System.Reflection.Assembly.LoadFile(lpProviderPath)
            lobjAssembly = System.Reflection.Assembly.LoadFrom(lstrProviderCandidatePath)
          Catch BadImageFormatEx As BadImageFormatException
            ' This would not be a provider in any case, just skip it
            Continue For
          Catch ex As Exception
            ApplicationLogging.LogException(ex, "ContentSource::GetProvider(lpTargetFolder, lpContentSource, bAutoConnect)")
            ' If we are unable to load the assembly then it is not valid.
            Throw New ApplicationException(Helper.FormatCallStack(ex), ex)
          End Try

          Try
            For Each lobjType As Type In lobjAssembly.GetTypes
              If lobjType.IsAbstract = False Then
                lobjProviderCandidate = lobjType.GetInterface(IPROVIDER)
                If lobjProviderCandidate IsNot Nothing Then
                  lobjProvider = lobjAssembly.CreateInstance(lobjType.FullName)
                  If lobjProvider.Name.Equals(lpContentSource.ProviderName) = False Then
                    Exit For
                  End If
                  lobjProvider.ContentSource = lpContentSource
                  Return lobjProvider
                End If
              End If
            Next
          Catch RefTypeLoadEx As Reflection.ReflectionTypeLoadException
            If RefTypeLoadEx.LoaderExceptions.Length > 0 Then
              Dim loaderExceptions As String = "Unable to load type from '" & lstrProviderCandidatePath & "'; "
              For Each loaderException As Exception In RefTypeLoadEx.LoaderExceptions
                loaderExceptions &= "Loader Exception: '" & loaderException.Message & "', "
              Next
              If loaderExceptions.EndsWith(", ") Then
                loaderExceptions = loaderExceptions.Remove(loaderExceptions.Length - 2)
              End If
              ApplicationLogging.WriteLogEntry(loaderExceptions, Reflection.MethodBase.GetCurrentMethod, TraceEventType.Error, 7457457)
            End If
            Continue For
          Catch Ex As Exception
            ' Just log it and skip it
            ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
            Continue For
          End Try
        Next

        Return Nothing

      Catch RefTypeLoadEx As Reflection.ReflectionTypeLoadException
        ApplicationLogging.LogException(RefTypeLoadEx, String.Format("ContentSource::GetProvider(lpTargetFolder:'{0}', lpContentSource, bAutoConnect)",
                                                          lstrProviderCandidatePath))
        If RefTypeLoadEx.LoaderExceptions.Length > 0 Then
          Dim loaderExceptions As String = "Unable to load type from '" & lstrProviderCandidatePath & "'; "
          For Each loaderException As Exception In RefTypeLoadEx.LoaderExceptions
            loaderExceptions &= "Loader Exception: '" & loaderException.Message & "', "
          Next
          If loaderExceptions.EndsWith(", ") Then
            loaderExceptions = loaderExceptions.Remove(loaderExceptions.Length - 2)
          End If
          Throw New ApplicationException(loaderExceptions, RefTypeLoadEx)
        Else
          Throw New ApplicationException(Helper.FormatCallStack(RefTypeLoadEx), RefTypeLoadEx)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Private Function FindProvider(Optional ByVal bAutoConnect As Boolean = True) As CProvider
    '  Try

    '    'LogSession.EnterMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

    '    'Dim lstrProviderPath As String = String.Empty
    '    Dim lobjProvider As CProvider = Nothing
    '    Dim lstrTargetFolder As String = String.Empty
    '    Dim lstrPreviousTargetFolder As String = String.Empty
    '    'Dim lobjCurrentConnectionSettings As Configuration.ConnectionSettings = Configuration.ConnectionSettings.GetCurrentSettings(False)
    '    Dim lobjCurrentConnectionSettings As Configuration.ConnectionSettings
    '    Dim lstrProviderLocation As String = String.Empty

    '    If Helper.CallStackContainsMethodName("AddContentSource") = True Or Helper.CallStackContainsMethodName("EditContentSource") = True Then
    '      lobjCurrentConnectionSettings = Configuration.ConnectionSettings.GetInstance(True)
    '    Else
    '      lobjCurrentConnectionSettings = Configuration.ConnectionSettings.Instance
    '    End If

    '    If lobjCurrentConnectionSettings Is Nothing Then
    '      lobjCurrentConnectionSettings = Configuration.ConnectionSettings.GetInstance(True)
    '    End If

    '    ' 1. Look in the provider catalog.
    '    'ApplicationLogging.WriteLogEntry(String.Format("1. Checking Provider Catalog for provider name '{0}'", Me.ProviderName), _
    '    '                                 Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information, 61231)
    '    Dim lobjProviderCatalog As ProviderCatalog = lobjCurrentConnectionSettings.ProviderCatalog
    '    'ApplicationLogging.WriteLogEntry(String.Format("Provider Catalog has {0} entries.", lobjProviderCatalog.Count), _
    '    '                                 Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information, 64220)
    '    If lobjProviderCatalog.Contains(Me.ProviderName) Then
    '      lobjProvider = GetProvider(lobjProviderCatalog(Me.ProviderName).ProviderPath, Me, bAutoConnect)
    '      Try
    '        lobjCurrentConnectionSettings.Serialize(lobjCurrentConnectionSettings.SettingsPath)
    '      Catch IoEx As IO.IOException
    '        If IoEx.Message.Contains("because it is being used by another process") Then
    '          ApplicationLogging.WriteLogEntry("Provider catalog currently being updated by another process.",
    '                                           Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 41001)
    '        Else
    '          ApplicationLogging.WriteLogEntry(String.Format("Skipping provider catalog update '{0}'", IoEx.Message),
    '                                         Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 41002)
    '        End If
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      End Try
    '      lstrProviderLocation = String.Format("1. Provider Path: '{0}' from Provider Catalog",
    '                                           lobjProvider.ProviderPath)
    '    Else
    '      ' 2. Try each of the providers listed in the 
    '      ' provider paths collection of the current settings file.
    '      'ApplicationLogging.WriteLogEntry(String.Format("2. Checking provider paths collection in current settings file '{0}'.",
    '      '                                               lobjCurrentConnectionSettings.SettingsPath),
    '      '                                             Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information, 61232)
    '      'LogSession.LogDebug(String.Format("2. Checking provider paths collection in current settings file '{0}'.",
    '      'lobjCurrentConnectionSettings.SettingsPath))

    '      For Each lstrProviderPath As String In lobjCurrentConnectionSettings.ProviderPaths
    '        lobjProvider = GetProvider(lstrProviderPath, Me, False)
    '        If lobjProvider IsNot Nothing Then
    '          ' Check to see if it is the correct provider
    '          If lobjProvider.Name = Me.ProviderName Then
    '            lstrProviderLocation = String.Format(
    '              "2. Provider Path: '{0}' in settings file: '{1}'",
    '              lstrProviderPath, lobjCurrentConnectionSettings.SettingsPath)
    '            Exit For
    '          Else
    '            lobjProvider = Nothing
    '          End If
    '        End If
    '      Next

    '      If lobjProvider Is Nothing Then
    '        ' 3. Look in the current application path
    '        lstrTargetFolder = My.Application.Info.DirectoryPath
    '        'ApplicationLogging.WriteLogEntry(String.Format("3. Checking current application path '{0}'.",
    '        '  lstrTargetFolder), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information, 61233)
    '        'LogSession.LogDebug(String.Format("3. Checking current application path '{0}'.",
    '        'lstrTargetFolder))
    '        lobjProvider = FindProvider(lstrTargetFolder, Me)

    '        If lobjProvider IsNot Nothing Then
    '          lstrProviderLocation = String.Format(
    '            "3. Provider Path: '{0}' in application path: '{1}'",
    '            lobjProvider.ProviderPath, lstrTargetFolder)
    '        Else
    '          ' 4. Look in the defined provider path, but only if 
    '          ' it is different from the current application path.
    '          lstrPreviousTargetFolder = lstrTargetFolder
    '          lstrTargetFolder = FileHelper.Instance.ProviderPath
    '          'ApplicationLogging.WriteLogEntry(String.Format("4. Checking defined provider path '{0}'.",
    '          '  lstrTargetFolder), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information, 61234)
    '          'LogSession.LogDebug(String.Format("4. Checking defined provider path '{0}'.",
    '          'lstrTargetFolder))

    '          If lstrPreviousTargetFolder.Equals(lstrTargetFolder) = False Then
    '            lobjProvider = FindProvider(lstrTargetFolder, Me)
    '          End If

    '          If lobjProvider IsNot Nothing Then
    '            lstrProviderLocation = String.Format(
    '              "4. Provider Path: '{0}' in defined provider path: '{1}'",
    '              lobjProvider.ProviderPath, lstrTargetFolder)
    '          Else
    '            ' 5. Look in the Cts root path, but only if 
    '            ' it is different from the defined provider path.
    '            lstrPreviousTargetFolder = lstrTargetFolder
    '            lstrTargetFolder = FileHelper.Instance.RootPath
    '            'ApplicationLogging.WriteLogEntry(String.Format("5. Checking Cts root path path '{0}'.",
    '            '  lstrTargetFolder), Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information, 61235)
    '            'LogSession.LogDebug(String.Format("5. Checking Cts root path path '{0}'.",
    '            'lstrTargetFolder))

    '            If lstrPreviousTargetFolder.Equals(lstrTargetFolder) = False Then
    '              lobjProvider = FindProvider(lstrTargetFolder, Me)
    '              lstrProviderLocation = String.Format(
    '                "5. Provider Path: '{0}' in Cts root path: '{1}'",
    '                lobjProvider.ProviderPath, lstrTargetFolder)
    '            End If
    '          End If
    '        End If
    '      End If
    '    End If

    '    If lobjProvider IsNot Nothing Then
    '      If Me.ConnectionString <> String.Empty Then
    '        lobjProvider.ConnectionString = Me.ConnectionString
    '      End If
    '      ' Add the provider to the catalog
    '      AddProviderToCatalog(lobjProvider)

    '      If lobjProvider.IsConnected = False AndAlso bAutoConnect = True Then
    '        lobjProvider.Connect(Me)
    '      End If

    '      '#If DEBUG Then
    '      'LogSession.LogDebug(String.Format("Found provider path for content source '{1}': {0}",
    '      'lstrProviderLocation, Me.Name),
    '      '  TraceEventType.Information, 35462)
    '      '#End If
    '      Return lobjProvider

    '    Else
    '      'ApplicationLogging.WriteLogEntry("Failed to find provider path", Reflection.MethodBase.GetCurrentMethod,
    '      '                                 TraceEventType.Warning, 61239)
    '      'LogSession.LogWarning("Failed to find provider path")
    '      Return Nothing
    '    End If

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  Finally
    '    'LogSession.LeaveMethod(Level.Debug, Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
    '  End Try
    'End Function

    'Private Sub AddProviderToCatalog(ByVal lpProvider As CProvider)
    '  Try

    '    Dim lobjCurrentConnectionSettings As Configuration.ConnectionSettings = Configuration.ConnectionSettings.Instance

    '    If lobjCurrentConnectionSettings Is Nothing Then
    '      lobjCurrentConnectionSettings = Configuration.ConnectionSettings.GetInstance(True)
    '    End If

    '    If lobjCurrentConnectionSettings.ProviderCatalog.Contains(lpProvider.Name) = False Then
    '      lobjCurrentConnectionSettings.ProviderCatalog.Add(lpProvider)
    '      Try
    '        lobjCurrentConnectionSettings.Serialize(lobjCurrentConnectionSettings.SettingsPath)
    '      Catch IoEx As IO.IOException
    '        If IoEx.Message.Contains("because it is being used by another process") Then
    '          ApplicationLogging.WriteLogEntry("Provider catalog currently being updated by another process.",
    '                                           Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 41001)
    '        Else
    '          ApplicationLogging.WriteLogEntry(String.Format("Skipping provider catalog update '{0}'", IoEx.Message),
    '                                         Reflection.MethodBase.GetCurrentMethod(), TraceEventType.Information, 41002)
    '        End If
    '      Catch ex As Exception
    '        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      End Try
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Private Function GetProvider(Optional ByVal lpAutoConnect As Boolean = True) As CProvider

      Try
        ' 'LogSession.EnterMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))

        '' Check to make sure that the ProviderPath has a value, if not we will throw an exception
        'If Me.ProviderPath Is Nothing Then
        '  Throw New InvalidPathException("The path is null", "") ';("Unable to get provider, ProviderPath is Nothing")
        'End If
        'If Me.ProviderPath.Length = 0 Then
        '  Throw New InvalidPathException(Me.ProviderPath, "Unable to get provider, the ProviderPath is not initialized")
        'End If
        Dim lobjProvider As CProvider = Nothing

        If String.IsNullOrEmpty(Me.ProviderPath) Then
          ' The ProviderPath is not available, we will try to 
          ' resolve the provider from one of the known paths.
          '
          ' 1. The provider catalog in the Settings.csf file.
          ' 2. The current application path.
          ' 3. The defined provider path.
          ' 4. The Cts root path.

          ''Try
          ''  lobjProvider = FindProvider(lpAutoConnect)
          ''Catch ex As Exception
          ''  ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 32981, lpAutoConnect)
          ''End Try

          Throw New InvalidOperationException("ProviderPath property not set...")

          'If lstrProviderPath Is Nothing OrElse lstrProviderPath.Length = 0 Then
          '  Throw New Exceptions.ProviderNotInitializedException(String.Format("Unable to locate provider path for content source '{0}'.", Me.Name))
          'End If
        Else
          Try
            lobjProvider = GetProvider(Me.ProviderPath, Me, lpAutoConnect)
          Catch ex As Exception
            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 32982, lpAutoConnect)
          End Try

        End If

        If lobjProvider IsNot Nothing AndAlso Not String.IsNullOrEmpty(Me.ConnectionString) Then
          lobjProvider.ConnectionString = Me.ConnectionString
        End If

        Return lobjProvider

      Catch pathex As InvalidPathException
        ' Wrap the exception in a ProviderNotInitializedException and attempt to insert the 
        ' ContentSource name in the message to simplify de-bugging the ContentSource connection string.
        Throw New Exceptions.ProviderNotInitializedException(pathex.Message.Replace("ContentSource does",
          String.Format("ContentSource '{0}' does", Me.Name)), Me, pathex)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 32984, lpAutoConnect)
        ' Re-throw the exception to the caller
        Throw
      Finally
        ' 'LogSession.LeaveMethod(Helper.GetMethodIdentifier(Reflection.MethodBase.GetCurrentMethod))
      End Try

    End Function

    Public Shared Function GetProvider(ByVal lpProviderPath As String,
                                       Optional ByVal lpContentSource As ContentSource = Nothing,
                                       Optional ByVal lpAutoConnect As Boolean = True) As Providers.CProvider
      Try

        Dim lobjProvider As Providers.CProvider

        ' Check to make sure that the ProviderPath has a value, if not we will throw an exception
        If lpProviderPath Is Nothing Then
          Throw New InvalidPathException("Unable to get provider, lpProviderPath is Nothing", "")
        End If
        If lpProviderPath.Length = 0 Then
          Throw New InvalidPathException("Unable to get provider, lpProviderPath is a zero length string", "")
        End If

        ' Check to make sure the path is valid, if not we will throw an exception
        If IO.File.Exists(lpProviderPath) = False Then
          Dim ioex As New IO.FileNotFoundException(String.Format("File does not exist: '{0}'.", lpProviderPath), lpProviderPath)
          Throw New InvalidPathException(String.Format("The provider path '{0}' specified for the ContentSource does not exist.", lpProviderPath), lpProviderPath, ioex)

        End If

        ' Validate the Provider

#If NET8_0_OR_GREATER Then
        If lpProviderPath.Contains("utilities", StringComparison.CurrentCultureIgnoreCase) Then
          ' This is a utility dll for a provider, do not attempt to load it.
          Return Nothing
        End If
#Else
        If lpProviderPath.ToLower.Contains("utilities") Then
          ' This is a utility dll for a provider, do not attempt to load it.
          Return Nothing
        End If
#End If

        If IsValidProvider(lpProviderPath) = True Then
          Debug.WriteLine(String.Format("The provider '{0}' is valid.", lpProviderPath))
        Else
          Dim lobjLastException As Exception = ExceptionTracker.LastException
          If lobjLastException IsNot Nothing Then
            Dim lstrErrorMessage As String = String.Format("The provider '{0}' is invalid: {1}|{2}.",
                                                          lpProviderPath, lobjLastException.GetType.Name, lobjLastException.Message)
            Debug.WriteLine(lstrErrorMessage)
            Throw New InvalidProviderException(lstrErrorMessage, lpProviderPath)
          Else
            Debug.WriteLine(String.Format("The provider '{0}' is invalid.", lpProviderPath))
            Throw New InvalidProviderException(lpProviderPath)
          End If
        End If

        lobjProvider = LoadProvider(lpProviderPath, lpContentSource, lpAutoConnect)

        Return lobjProvider

        'Try
        '  'lobjAssembly = System.Reflection.Assembly.LoadFile(lpProviderPath)
        '  lobjAssembly = System.Reflection.Assembly.LoadFrom(lpProviderPath)
        'Catch ex As Exception
        '  ApplicationLogging.LogException(ex, "ContentSource::GetProvider(lpProviderPath, Optional lpContentSource)")
        '  ' If we are unable to load the assembly then it is not valid.
        '  Throw New ApplicationException(Helper.FormatCallStack(ex), ex)
        'End Try

        'For Each lobjType As Type In lobjAssembly.GetTypes
        '  lobjProviderCandidate = lobjType.GetInterface(IPROVIDER)
        '  If Not lobjProviderCandidate Is Nothing Then
        '    lobjProvider = lobjAssembly.CreateInstance(lobjType.FullName)
        '    lobjProvider.ContentSource = lpContentSource
        '    If (bAutoConnect) Then
        '      If lobjProvider.IsConnected = False AndAlso lpContentSource IsNot Nothing Then
        '        lobjProvider.Connect(lpContentSource)
        '      End If
        '    End If
        '    Return lobjProvider
        '  End If
        'Next

        'Return Nothing

      Catch ex As Reflection.ReflectionTypeLoadException
        ApplicationLogging.LogException(ex, String.Format("ContentSource::GetProvider(lpProviderPath:'{0}', Optional lpContentSource)", lpProviderPath))
        If ex.LoaderExceptions.Length > 0 Then
          Dim loaderExceptions As String = "Unable to load type from '" & lpProviderPath & "'; "
          For Each loaderException As Exception In ex.LoaderExceptions
            loaderExceptions &= "Loader Exception: '" & loaderException.Message & "', "
          Next
          If loaderExceptions.EndsWith(", ") Then
            loaderExceptions = loaderExceptions.Remove(loaderExceptions.Length - 2)
          End If
          Throw New ApplicationException(loaderExceptions, ex)
        Else
          Throw New ApplicationException(Helper.FormatCallStack(ex), ex)
        End If
      Catch ex As Exception
        If (ex.InnerException IsNot Nothing) AndAlso (ex.InnerException.GetType.Name = "LicenseException") Then
          ApplicationLogging.LogException(ex.InnerException, String.Format("ContentSource::GetProvider(lpProviderPath:'{0}', Optional lpContentSource)", lpProviderPath))
          ' Re-throw the exception to the caller
          Throw ex.InnerException
        End If
        ApplicationLogging.LogException(ex, String.Format("ContentSource::GetProvider(lpProviderPath:'{0}', Optional lpContentSource)", lpProviderPath))
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Validates the provider dll by determining if it 
    ''' contains a type that implements the IProvider interface.
    ''' </summary>
    ''' <param name="lpProviderPath">The fully qualified path to the dll to check.</param>
    ''' <returns>True if IProvider is implemented, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Shared Function IsValidProvider(ByVal lpProviderPath As String) As Boolean

      Try

        Dim lobjAssembly As System.Reflection.Assembly
        Dim lobjProviderCandidate As Type

        Try
          'lobjAssembly = System.Reflection.Assembly.LoadFile(lpProviderPath)
          lobjAssembly = System.Reflection.Assembly.LoadFrom(lpProviderPath)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 64328, lpProviderPath)
          ' If we are unable to load the assembly then it is not valid.
          Return False
        End Try

        For Each lobjType As Type In lobjAssembly.GetTypes
          lobjProviderCandidate = lobjType.GetInterface(IPROVIDER)
          If lobjProviderCandidate IsNot Nothing Then
            Return True
          End If
        Next

        Return False
      Catch Reflex As Reflection.ReflectionTypeLoadException
        ApplicationLogging.LogException(Reflex, Reflection.MethodBase.GetCurrentMethod, 64250, lpProviderPath)
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 64251, lpProviderPath)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Shared Function LoadProvider(ByVal lpProviderPath As String, ByVal lpContentSource As ContentSource, ByVal lpAutoConnect As Boolean) As IProvider
      Try

        Dim lobjAssembly As System.Reflection.Assembly
        Dim lobjProviderCandidate As Type
        Dim lobjProvider As Providers.CProvider

        Try
          'lobjAssembly = System.Reflection.Assembly.LoadFile(lpProviderPath)
          lobjAssembly = System.Reflection.Assembly.LoadFrom(lpProviderPath)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 64329, lpProviderPath)
          ' If we are unable to load the assembly then it is not valid.
          Return Nothing
        End Try

        Dim lobjTypes As Type() = lobjAssembly.GetTypes

        For Each lobjType As Type In lobjTypes
          lobjProviderCandidate = lobjType.GetInterface(IPROVIDER)
          If lobjProviderCandidate IsNot Nothing Then
            lobjProvider = lobjAssembly.CreateInstance(lobjType.FullName)
            lobjProvider.ContentSource = lpContentSource
            If (lpAutoConnect) Then
              If lobjProvider.IsConnected = False AndAlso lpContentSource IsNot Nothing Then
                lobjProvider.Connect(lpContentSource)
              End If
            End If
            Return lobjProvider
          End If
        Next

        Return Nothing

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 64252, lpProviderPath)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Public Shared Function GetProvider(ByVal lpProviderPath As String) As Ecmg.Cts.Providers.CProvider
    '  Try
    '    Dim lobjAssembly As System.Reflection.Assembly
    '    Dim lobjProviderCandidate As Type
    '    Dim lobjProvider As Ecmg.Cts.Providers.CProvider

    '    ' Validate the Provider
    '    Try
    '      'lobjAssembly = System.Reflection.Assembly.LoadFile(lpProviderPath)
    '      lobjAssembly = System.Reflection.Assembly.LoadFrom(lpProviderPath)
    '    Catch ex As Exception
    '      ' If we are unable to load the assembly then it is not valid.
    '      Throw New ApplicationException(Helper.FormatCallStack(ex), ex)
    '    End Try

    '    For Each lobjType As Type In lobjAssembly.GetTypes
    '      lobjProviderCandidate = lobjType.GetInterface(IPROVIDER)
    '      If Not lobjProviderCandidate Is Nothing Then
    '        lobjProvider = lobjAssembly.CreateInstance(lobjType.FullName)
    '        If lobjProvider.IsConnected = False Then
    '          lobjProvider.Connect(Me)
    '        End If
    '        Return lobjProvider
    '      End If
    '    Next

    '    Return Nothing

    '  Catch ex As Reflection.ReflectionTypeLoadException
    '    If ex.LoaderExceptions.Length > 0 Then
    '      Dim loaderExceptions As String = "Unable to load type from '" & lpProviderPath & "'; "
    '      For Each loaderException As Exception In ex.LoaderExceptions
    '        loaderExceptions &= "Loader Exception: '" & loaderException.Message & "', "
    '      Next
    '      If loaderExceptions.EndsWith(", ") Then
    '        loaderExceptions = loaderExceptions.Remove(loaderExceptions.Length - 2)
    '      End If
    '      Throw New ApplicationException(loaderExceptions, ex)
    '    Else
    '      Throw New ApplicationException(Helper.FormatCallStack(ex), ex)
    '    End If
    '  Catch ex As Exception
    '    Throw New ApplicationException(Helper.FormatCallStack(ex), ex)
    '  End Try

    'End Function

    'Private Sub CheckProviderLicense()
    '  Try
    '    If Provider Is Nothing Then
    '      If ExceptionTracker.LastException IsNot Nothing Then
    '        Throw New ProviderNotInitializedException(
    '          String.Format("Unable to check provider license for content source '{0}: {1} ({2})'",
    '                        Me.ProviderName, Me.Name, ExceptionTracker.LastException.Message), Me, ExceptionTracker.LastException)
    '      Else
    '        Throw New ProviderNotInitializedException(Me)
    '      End If
    '    End If
    '    If Provider.HasValidLicense = False Then
    '      If CtsLicense.FileExists Then
    '        ' Throw New InvalidLicenseException(CtsLicense.Instance.NativeLicense)
    '        Throw New InvalidLicenseException(LogicNP.CryptoLicensing.LicenseStatus.InValid)
    '      Else
    '        Throw New LicenseFileNotFoundException
    '      End If
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, 64248)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Private Function GetIsActiveValue() As Boolean
      Try
        If Properties.Contains("Active") Then
          Dim lobjActiveProperty As ProviderProperty = Properties("Active")
          If lobjActiveProperty IsNot Nothing Then
            If lobjActiveProperty.PropertyValue IsNot Nothing Then
              If TypeOf lobjActiveProperty.PropertyValue Is Boolean Then
                Return lobjActiveProperty.PropertyValue
              ElseIf TypeOf lobjActiveProperty.PropertyValue Is String Then

                Dim lblnIsActive As Boolean = False
                Dim lblnIsValidBool As Boolean = Boolean.TryParse(lobjActiveProperty.PropertyValue, lblnIsActive)
                If lblnIsValidBool = True Then
                  Return lblnIsActive
                Else
                  Return True
                End If

              Else
                Return True
              End If
            End If
            Return lobjActiveProperty.PropertyValue
          Else
            Return True
          End If
        Else
          Return True
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Private Sub InitializeLogSession()
    '  Try
    '    'mobjLogSession = SiAuto.Si.AddSession(Me.GetType.Name)
    '    'mobjLogSession.Color = Drawing.Color.LightGray
    '    mobjLogSession = ApplicationLogging.InitializeLogSession(Me.GetType.Name, Drawing.Color.LightGray)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    'Private Sub FinalizeLogSession()
    '  Try
    '    'LogSession.LogMessage("ContentSource: {0} disposing on thread({1})", Me.Name, Threading.Thread.CurrentThread.ManagedThreadId)
    '    'SiAuto.Si.DeleteSession(LogSession)
    '    ApplicationLogging.FinalizeLogSession(LogSession)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Private Sub InitializePropertiesFromConnectionString()

      Try
        'LogSession.LogVerbose("Starting ContentSource.InitializePropertiesFromConnectionString")
        mobjProviderProperties = GetPropertiesFromConnectionString(mstrConnectionString)

        ' Try to get the ContentSource name from the connection string
        Dim lstrName As String = Helper.GetInfoFromString(mstrConnectionString, "Name")
        Me.Name = lstrName

        'mstrConnectionString = Provider(False).GenerateConnectionString(Me.Name, Me.Properties)
        Active = GetIsActiveValue()

        'LogSession.LogVerbose("Completed ContentSource({0}).InitializePropertiesFromConnectionString", Me.Name)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "ContentSource::GetPropertiesFromConnectionString")
        ' Rew-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Shared Function GetNameFromConnectionString(ByVal lpConnectionString As String) As String
      Try
        Dim lobjProviderProperties As ProviderProperties = GetPropertiesFromConnectionString(lpConnectionString)

        Dim lstrName As String = (From lobjProperty In lobjProviderProperties Where
               lobjProperty.PropertyName = "Name" Select lobjProperty.PropertyValue).First

        Return lstrName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetProviderNameFromConnectionString(ByVal lpConnectionString As String) As String
      Try
        Dim lobjProviderProperties As ProviderProperties = GetPropertiesFromConnectionString(lpConnectionString)

        Dim lstrName As String = (From lobjProperty In lobjProviderProperties Where
               lobjProperty.PropertyName = "Provider" Select lobjProperty.PropertyValue).First

        Return lstrName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetProviderPathFromConnectionString(ByVal lpConnectionString As String) As String
      Try
        Dim lobjProviderProperties As ProviderProperties = GetPropertiesFromConnectionString(lpConnectionString)

        Dim lstrName As String = (From lobjProperty In lobjProviderProperties Where
               lobjProperty.PropertyName = "ProviderPath" Select lobjProperty.PropertyValue).FirstOrDefault

        Return lstrName

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shared Function GetPropertiesFromConnectionString(ByVal lpConnectionString As String,
                                                                     Optional ByVal lpDelimiter As String = ";", Optional ByVal lpNestedConnectionStringDelimiter As String = "'") As ProviderProperties

      Try

        ' We need to first look for any possible nested connection strings
        ' Look for strings inside curly braces { }
        Dim lstrNestedStringPattern As String = "(?<={).*?(?=})"
        Dim lobjNestedStrings As System.Text.RegularExpressions.MatchCollection = Regex.Matches(lpConnectionString, lstrNestedStringPattern)
        Dim lstrCompleteConnectionString As String = lpConnectionString
        Dim lstrNestedString As String = ""
        Dim iNestedCounter As Integer = 0

        If (lobjNestedStrings IsNot Nothing) AndAlso (lobjNestedStrings.Count > 0) Then
          ' We have at least one nested connection string
          'lstrNestedString = lobjNestedStrings(0).Value

          For Each lstrString As Match In lobjNestedStrings
            ' Remove the nested string from the complete connection string
            lstrCompleteConnectionString = lstrCompleteConnectionString.Replace(lstrString.Value, "")
          Next

        End If

        Dim lstrNameValuePairs() As String = lstrCompleteConnectionString.Split(lpDelimiter)
        Dim lstrNameValuePair() As String
        Dim lblnInfoFound As Boolean = False
        Dim lobjProviderProperty As ProviderProperty
        Dim lobjProviderProperties As New ProviderProperties

        If lstrCompleteConnectionString = "" Then
          'Return Nothing
          Return lobjProviderProperties
        End If

        For Each lstrPair As String In lstrNameValuePairs

          'debug.writeline(lstrPair)
          If (lstrPair.Length = 0) Then
            Continue For
          End If

          lstrNameValuePair = lstrPair.Trim.Split("=")
          If (lstrNameValuePair.Length > 1) AndAlso (lstrNameValuePair(1) = "{}") Then
            ' Replace the {} stub with the actual nested string
            lstrNameValuePair(1) = lobjNestedStrings(iNestedCounter).Value
            iNestedCounter += 1
          End If


          lobjProviderProperty = New ProviderProperty(lstrNameValuePair(0), lstrNameValuePair(1))
          lobjProviderProperties.Add(lobjProviderProperty)
        Next

        Return lobjProviderProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("ContentSource::GetPropertiesFromConnectionString('{0}', '{1}', '{2}')", lpConnectionString, lpDelimiter, lpNestedConnectionStringDelimiter))
        Throw New Exception("Unable to get properties from connection string", ex)
      End Try

    End Function

    ''' <summary>
    '''     Returns an updated connection string with the new property value
    ''' </summary>
    ''' <param name="lpConnectionString" type="String">
    '''     <para>
    '''         The connection string to update.
    '''     </para>
    ''' </param>
    ''' <param name="lpPropertyName" type="String">
    '''     <para>
    '''         The target property to update.
    '''     </para>
    ''' </param>
    ''' <param name="lpNewPropertyValue" type="String">
    '''     <para>
    '''         The new value to set for the target property.
    '''     </para>
    ''' </param>
    ''' <returns>
    '''     A String value...
    ''' </returns>
    Public Shared Function ChangeConnectionStringValue(ByVal lpConnectionString As String,
                                                       ByVal lpPropertyName As String,
                                                       ByVal lpNewPropertyValue As String) As String
      Try
        Dim lstrReturnValue As String = lpConnectionString

        Dim lobjProperties As ProviderProperties = GetPropertiesFromConnectionString(lpConnectionString)
        If lobjProperties.Contains(lpPropertyName) = False Then
          Throw New Exceptions.PropertyDoesNotExistException(lpPropertyName)
        End If

        lobjProperties.Item(lpPropertyName).Value = lpNewPropertyValue
        Dim lstrContentSourceName As String = GetNameFromConnectionString(lpConnectionString)
        Dim lstrProviderName As String = GetProviderNameFromConnectionString(lpConnectionString)
        Dim lstrProviderPath As String = GetProviderPathFromConnectionString(lpConnectionString)

        If lobjProperties.Contains("ProviderPath") Then
          lstrReturnValue = CProvider.GenerateConnectionStringShared(lstrContentSourceName,
                                                   lobjProperties,
                                                   lstrProviderName,
                                                   lstrProviderPath)
        Else
          lstrReturnValue = CProvider.GenerateConnectionStringShared(lstrContentSourceName,
                                                   lobjProperties,
                                                   lstrProviderName,
                                                   String.Empty)
        End If

        Return lstrReturnValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub MobjProvider_ConnectionStateChanged(sender As Object, ByRef e As Arguments.ConnectionStateChangedEventArgs) Handles MobjProvider.ConnectionStateChanged
      Try
        State = e.CurrentState
        RaiseEvent ConnectionStateChanged(Me, e)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    Public ReadOnly Property IsDisposed() As Boolean Implements IRepositoryConnection.IsDisposed
      Get
        Return disposedValue
      End Get
    End Property

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            ' DISPOSETODO: dispose managed state (managed objects).
            'FinalizeLogSession()
            mstrImportPath = String.Empty
            mstrExportPath = String.Empty
            mobjProviderProperties = New ProviderProperties
            mblnActive = False
            mstrName = String.Empty
            mstrDescription = String.Empty
            mstrProviderName = String.Empty
            mstrSerializationPath = Nothing
            mstrConnectionString = String.Empty
            mstrProviderPath = String.Empty
            MobjProvider.Dispose()
            MobjProvider = Nothing
            menuState = ProviderConnectionState.Disconnected
            mobjSecurityToken = Nothing
            mobjSelectedFolder = Nothing
          End If

          ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
          ' DISPOSETODO: set large fields to null.
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        Me.disposedValue = True
      End Try
    End Sub

    ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub

    'Public Overrides Function GetHashCode() As Integer
    '  Throw New NotImplementedException()
    'End Function

    Public Shared Operator =(left As ContentSource, right As ContentSource) As Boolean
      If left Is Nothing Then
        Return right Is Nothing
      End If

      Return left.Equals(right)
    End Operator

    Public Shared Operator <>(left As ContentSource, right As ContentSource) As Boolean
      Return Not left = right
    End Operator

    Public Shared Operator <(left As ContentSource, right As ContentSource) As Boolean
      If left Is Nothing Then
        Return right IsNot Nothing
      Else
        Return left.CompareTo(right) < 0
      End If
    End Operator

    Public Shared Operator <=(left As ContentSource, right As ContentSource) As Boolean
      Return left Is Nothing OrElse left.CompareTo(right) <= 0
    End Operator

    Public Shared Operator >(left As ContentSource, right As ContentSource) As Boolean
      Return left IsNot Nothing AndAlso left.CompareTo(right) > 0
    End Operator

    Public Shared Operator >=(left As ContentSource, right As ContentSource) As Boolean
      Return If(left Is Nothing, right Is Nothing, left.CompareTo(right) >= 0)
    End Operator
#End Region

  End Class

End Namespace