'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Reflection
Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Arguments
Imports Documents.Core
Imports Documents.Core.ChoiceLists
Imports Documents.SerializationUtilities
Imports Documents.Utilities

#End Region

Namespace Providers

  ''' <remarks>This object is analogous to an ODBC provider.</remarks>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class CProvider
    Inherits Core.CtsObject
    Implements IProvider
    Implements IAuthorization
    Implements ISerialize
    Implements IXmlSerializable

#Region "Class Constants"

    Public Const ACTION_SET_SYSTEM_PROPERTIES As String = "SetSystemProperties"
    Public Const ALLOW_ZERO_LENGTH_CONTENT As String = "AllowZeroLengthContent"
    Public Const EXPORT_AS_ARCHIVE As String = "ExportAsArchive"
    Public Const CTS_DOCS_PATH_WILDCARD As String = "%CtsDocsPath%"
    Protected Const CTS_DOCS_PATH_REPLACEMENT As String = CTS_DOCS_PATH_WILDCARD & "\"
    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    Public Const PROVIDER_INFORMATION_FILE_EXTENSION As String = "pif"

    Protected Friend Const EXPORT_PATH As String = "ExportPath"
    Protected Friend Const IMPORT_PATH As String = "ImportPath"
    Protected Friend Const ENFORCE_CLASSIFICATION_COMPLIANCE As String = "EnforceClassificationCompliance"
    Protected Friend Const LOG_INVALID_PROPERTY_REMOVALS As String = "LogInvalidPropertyRemovals"

    Protected Friend Const USER As String = "UserName"
    Protected Friend Const PWD As String = "Password"

    Protected Friend Const EXPORT_VERSION_SCOPE As String = "ExportVersionScope"
    Protected Friend Const EXPORT_MAX_VERSIONS As String = "ExportMaxVersions"

#End Region

#Region "Class Variables"

    'Private mstrConnectionString As String
    Private menuProviderClass As New ProviderClass
    Private mobjProviderSystem As New ProviderSystem
    Private mstrSerializationPath As String
    Private mstrExportPath As String
    'Private mstrImportPath As String
    'Private mobjProviderConfiguration As ProviderConfiguration
    Private WithEvents MobjProviderProperties As New ProviderProperties
    Private WithEvents MobjBackgroundWorker As New BackgroundWorker
    Private mblnIsConnected As Boolean
    Private mblnIsInitialized As Boolean
    'Private mobjSearch As CSearch
    'Private mobjImageSet As New ImageSet
    Private mintMaxContentCount As Integer = -1
    Private mobjProviderInformation As ProviderInformation
    Private mobjContentSource As ContentSource
    Private mobjActionProperties As New ActionProperties
    Private mblnHasValidLicense As Boolean
    Private mstrLicenseFailureReason As String
    Protected mblnEnforceClassificationCompliance As Boolean = True
    Private mblnLogInvalidPropertyRemovals As Nullable(Of Boolean)
    Private mblnClassificationPropertiesInitialized As Boolean
    Private menuState As ProviderConnectionState = ProviderConnectionState.Disconnected
    Private mblnAllowZeroLengthContent As Nullable(Of Boolean)
    ' For IExplorer
    Private mobjSelectedFolder As IFolder
    Private mobjTag As Object

#End Region

#Region "Public Events"

    Public Event ConnectionStateChanged As ConnectionStateChangedEventHandler Implements IProvider.ConnectionStateChanged

#Region "Background Worker Events"

    Event BackgroundWorker_Disposed As EventHandler Implements IProvider.BackgroundWorker_Disposed
    Event BackgroundWorker_DoWork As System.ComponentModel.DoWorkEventHandler Implements IProvider.BackgroundWorker_DoWork
    Event BackgroundWorker_ProgressChanged As System.ComponentModel.ProgressChangedEventHandler Implements IProvider.BackgroundWorker_ProgressChanged
    Event BackgroundWorker_RunWorkerCompleted As System.ComponentModel.RunWorkerCompletedEventHandler Implements IProvider.BackgroundWorker_RunWorkerCompleted


#End Region

#Region "ProviderProperty Events"

    Public Event ProviderProperty_ValueChanged As ProviderPropertyValueChangedEventHandler Implements IProvider.ProviderProperty_ValueChanged
    Public Event ProviderPropertyChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs)

#End Region

#End Region

#Region "Public Delegates"

    Public Delegate Sub UpdateSystemPropertiesCallback(ByVal Args As UpdateSystemPropertiesArgs)

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
      Try
        InitializeBaseInformation()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal ConnectionString As String)
      MyBase.New(ConnectionString)
      Try
        InitializeBaseInformation()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' <Removed by: Ernie Bahr at: 9/24/2012-09:43:01 on machine: ERNIEBAHR-THINK>
    ' Used the old license model
    '     Public Sub ValidateLicense()
    '       Try
    '         Dim rsn As New LicenseFailureReason()
    '         Dim lic As New ProviderLicense(Me, FileHelper.Instance.LicensePath & "EcmgProvider.slf")
    '         rsn = lic.CanContinue()
    '         mobjLicenseFailureReason = rsn
    '         If Not rsn.Reason = LicenseFailureReason.eReason.ok Then
    '           'Throw New License.LicenseException(rsn)
    '           mblnHasValidLicense = False
    '         Else
    '           mblnHasValidLicense = True
    '         End If
    '       Catch LicEx As License.LicenseException
    '         ApplicationLogging.WriteLogEntry(String.Format("Provider '{0}' could not be instantiated: '{1}'", Me.Name, LicEx.Message), TraceEventType.Error)
    '         '  Re-throw the exception to the caller
    '         Throw
    '       Catch ex As Exception
    '         ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '         '  Re-throw the exception to the caller
    '         Throw
    '       End Try
    '     End Sub
    ' </Removed by: Ernie Bahr at: 9/24/2012-09:43:01 on machine: ERNIEBAHR-THINK>

    '' <Added by: Ernie Bahr at: 9/24/2012-09:43:13 on machine: ERNIEBAHR-THINK>
    '' Uses new license model
    'Public Sub ValidateLicense()
    '  Try
    '    'Dim lobjLicense As CtsLicense = CtsLicense.Instance
    '    'lobjLicense.IsFeatureLicensed(
    '    ' If CtsLicense.IsFeatureLicensed(Me.Feature) = True Then
    '    If LicenseRegister.IsFeatureLicensed(Me.Feature) = True Then
    '      mblnHasValidLicense = True
    '    Else
    '      mblnHasValidLicense = False
    '      ' Throw New Exceptions.ProviderNotLicensedException(CtsLicense.Instance.NativeLicense, Me.Feature)
    '      Throw New Exceptions.ProviderNotLicensedException(Me.Feature)
    '    End If
    '  Catch LicenseEx As Exceptions.LicenseException
    '    ApplicationLogging.WriteLogEntry(LicenseEx.Message, Reflection.MethodBase.GetCurrentMethod,
    '                                     TraceEventType.Warning, 64900)
    '    mstrLicenseFailureReason = LicenseEx.Message
    '    mblnHasValidLicense = False
    '    ' Re-throw the exception to the caller
    '    Throw
    '  Catch Ex As Exception
    '    ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub
    ' </Added by: Ernie Bahr at: 9/24/2012-09:43:13 on machine: ERNIEBAHR-THINK>

    'Public Sub New(ByVal lpContentSource As ContentSource)
    '  MyBase.New(lpContentSource.ConnectionString)
    '  InitializeProvider(lpContentSource)
    'End Sub

    ''' <summary>
    ''' Factory method for instantiating a provider from its dll path
    ''' </summary>
    ''' <param name="lpProviderPath">Fully qualified file name of provider file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Create(ByVal lpProviderPath As String) As CProvider

      Try
        Dim lobjAssembly As System.Reflection.Assembly
        Dim lobjProviderCandidate As Type
        Dim lobjProvider As CProvider

        lobjAssembly = Reflection.Assembly.LoadFrom(lpProviderPath)

        For Each lobjType As Type In lobjAssembly.GetTypes
          lobjProviderCandidate = lobjType.GetInterface("IProvider")
          If lobjProviderCandidate IsNot Nothing Then
            lobjProvider = lobjAssembly.CreateInstance(lobjType.FullName)
            Return lobjProvider
          End If
        Next

        Return Nothing

      Catch TargetEx As Reflection.TargetInvocationException
        If TargetEx.InnerException IsNot Nothing Then
          Throw TargetEx.InnerException
        Else
          ApplicationLogging.LogException(TargetEx, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "Public Properties"

    Public Overrides Property UserName As String Implements IProvider.UserName
      Get
        Return MyBase.UserName
      End Get
      Set(value As String)
        MyBase.UserName = value
        SetStringPropertyValue(USER, value)
      End Set
    End Property

    Public Shadows Property ProviderPassword As String Implements IProvider.Password
      Protected Get
        Debug.WriteLine(Helper.FormatCallStack)
        Return MyBase.Password
        'Return String.Empty
      End Get
      Set(value As String)
        MyBase.Password = value
        SetStringPropertyValue(PWD, value)
      End Set
    End Property

    Public ReadOnly Property HasValidLicense() As Boolean Implements IProvider.HasValidLicense
      Get
        Return mblnHasValidLicense
      End Get
    End Property

    Public ReadOnly Property LicenseFailureReason() As String Implements IProvider.LicenseFailureReason
      Get
        Return mstrLicenseFailureReason
      End Get
    End Property

    Public Overridable ReadOnly Property TokenRequired As Boolean Implements IProvider.TokenRequired
      Get
        Return False
      End Get
    End Property

    ''' <summary>
    ''' Specifies the maximum number of content items to initialize folders with
    ''' -1 = get all content items
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property MaxContentCount() As Integer
      Get
        Return mintMaxContentCount
      End Get
      Set(ByVal value As Integer)
        mintMaxContentCount = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not the provider is currently in a connected state.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    Public Property IsConnected() As Boolean
      Get
        Return mblnIsConnected
      End Get
      Protected Set(ByVal value As Boolean)
        Try
          mblnIsConnected = value
          If value = True Then
            SetState(ProviderConnectionState.Connected)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ' <Added by: Ernie at: 3/1/2012-8:21:27 AM on machine: ERNIE-M4400>
    ' Added in response to FogBugz case 245 from Duke Energy.
    ' They had cases where they had intentionally stored zero length files in FileNet Content Services.
    ' Add the 'AllowZeroLengthContent' property
    ''' <summary>
    ''' Gets or sets a value determining whether or not to allow 
    ''' zero length content files (0 bytes) to be exported 
    ''' without raising a ZeroLengthContentException.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    Public Property AllowZeroLengthContent As Boolean
      Get
        Try
          If mblnAllowZeroLengthContent Is Nothing Then
            mblnAllowZeroLengthContent = GetBooleanPropertyValue(ALLOW_ZERO_LENGTH_CONTENT)
          End If
          Return mblnAllowZeroLengthContent
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Boolean)
        Try
          SetBooleanPropertyValue(ALLOW_ZERO_LENGTH_CONTENT, value)
          mblnAllowZeroLengthContent = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property
    ' </Added by: Ernie at: 3/1/2012-8:21:27 AM on machine: ERNIE-M4400>

    Public Property ExportAsArchive As Boolean
      Get
        Try
          Return GetBooleanPropertyValue(EXPORT_AS_ARCHIVE)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Boolean)
        Try
          SetBooleanPropertyValue(EXPORT_AS_ARCHIVE, value)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets the current connection status of the provider.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IsConnectedEx() As Boolean Implements IProvider.IsConnected
      Get
        Return mblnIsConnected
      End Get
    End Property

    ''' <summary>
    ''' Gets the current initialization status of the provider.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>A value of true does not necessarily mean that 
    ''' the provider is connected to the repository, only that 
    ''' it has been initialized and all of its proeprties set.</remarks>
    Public ReadOnly Property IsInitialized() As Boolean Implements IProvider.IsInitialized
      Get
        Return mblnIsInitialized
      End Get
    End Property

    '''' <summary>
    '''' Gets or Sets the set of images to be used when for the provider when displayed in a treeview or other control
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Property ImageSet() As ImageSet
    '  Get
    '    Return mobjImageSet
    '  End Get
    '  Set(ByVal value As ImageSet)
    '    mobjImageSet = value
    '  End Set
    'End Property

    ''' <summary>
    ''' Gets a collection of available ActionProperties for this provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ActionProperties() As ActionProperties Implements IProvider.ActionProperties
      Get
        Return mobjActionProperties
      End Get
    End Property

    ''' <summary>
    ''' Gets the requested ActionProperty for this provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ActionProperties(ByVal lpName As String) As ActionProperty
      Get
        Try
          Return mobjActionProperties(lpName)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the collection of provider properties
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProviderProperties() As ProviderProperties Implements IProvider.ProviderProperties
      Get
        Return MobjProviderProperties
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the default path to which the provider will export documents.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property ExportPath() As String
      Get
        Try
          If Right(mstrExportPath, 1) <> "\" Then
            Return mstrExportPath & "\"
          Else
            Return mstrExportPath
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        mstrExportPath = value
      End Set
    End Property

    Public ReadOnly Property LogInvalidPropertyRemovals As Boolean
      Get
        Try
          If mblnLogInvalidPropertyRemovals.HasValue Then
            Return mblnLogInvalidPropertyRemovals
          Else
            Return False
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property ResolvedExportPath As String
      Get
        Return ExportPath.Replace(CTS_DOCS_PATH_REPLACEMENT, FileHelper.Instance.CtsDocsPath)
      End Get
    End Property

    ' <Removed by: Ernie at: 9/29/2014-11:16:47 AM on machine: ERNIE-THINK>
    '     ''' <summary>
    '     ''' Gets or sets the default path from which the provider will import documents.
    '     ''' </summary>
    '     ''' <value></value>
    '     ''' <returns></returns>
    '     ''' <remarks></remarks>
    '     Public Overridable Property ImportPath() As String
    '       Get
    '         Try
    '           If Right(mstrImportPath, 1) <> "\" Then
    '             Return mstrImportPath & "\"
    '           Else
    '             Return mstrImportPath
    '           End If
    '         Catch ex As Exception
    '           ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '           ' Re-throw the exception to the caller
    '           Throw
    '         End Try
    '       End Get
    '       Set(ByVal Value As String)
    '         mstrImportPath = Value
    '       End Set
    '     End Property
    ' </Removed by: Ernie at: 9/29/2014-11:16:47 AM on machine: ERNIE-THINK>

    'Public ReadOnly Property ResolvedImportPath As String
    '  Get
    '    Return ImportPath.Replace(CTS_DOCS_PATH_REPLACEMENT, FileHelper.Instance.CtsDocsPath)
    '  End Get
    'End Property

    ''' <summary>
    ''' Gets the provider system.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The values initialized for the provider system are unique for each provider implementation.</remarks>
    Public Overridable ReadOnly Property ProviderSystem() As ProviderSystem
      Get
        Return mobjProviderSystem
        'Throw New NotImplementedException("'ProviderSystem' must be implemented in an inherited class")
      End Get
    End Property

    ''' <summary>
    ''' Gets the Provider Class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property ClassType() As ProviderClass
      Get
        Return menuProviderClass
      End Get
    End Property

    'Public Overridable ReadOnly Property Feature As FeatureEnum Implements IProvider.Feature
    '  Get
    '    Try
    '      Return FeatureEnum.FeatureNotSpecified
    '    Catch Ex As Exception
    '      ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property

    ''' <summary>
    ''' Gets the name of the provider.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Overloads ReadOnly Property Name() As String Implements IProvider.Name
      Get
        Try
          If ProviderSystem IsNot Nothing AndAlso ProviderSystem.Name.Length > 0 Then
            Return ProviderSystem.Name
          Else
            Return String.Empty
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::Get_Name", Me.GetType.Name))
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets a ProviderInformation object 
    ''' describing the current provider
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property Information() As ProviderInformation Implements IProvider.Information
      Get
        Try
          If mobjProviderInformation Is Nothing Then
            mobjProviderInformation = New ProviderInformation(Me)
          End If
          Return mobjProviderInformation
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the Search object utilized by the class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This is implemented by each provider.</remarks>
    Public MustOverride ReadOnly Property Search() As ISearch Implements IProvider.Search

    ''' <summary>
    ''' Gets the fully qualified path of the dll implementing the provider.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProviderPath() As String
      Get
        Return GetProviderPath()
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets parent content source reference
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    Public Property ContentSource() As ContentSource Implements IProvider.ContentSource
      Get
        Return mobjContentSource
      End Get
      Set(ByVal value As ContentSource)
        mobjContentSource = value
      End Set
    End Property

    Public Property SelectedFolder() As IFolder Implements IProvider.SelectedFolder
      Get
        Return mobjSelectedFolder
      End Get
      'Protected Friend Set(ByVal value As IFolder)
      Set(ByVal value As IFolder)
        mobjSelectedFolder = value
      End Set
    End Property

    Public ReadOnly Property State As ProviderConnectionState Implements IProvider.State
      Get
        Return menuState
      End Get
    End Property

    Public Property Tag As Object Implements IProvider.Tag
      Get
        Return mobjTag
      End Get
      Set(value As Object)
        mobjTag = value
      End Set
    End Property

    Protected Shared Function FolderExists(ByVal sender As IFolderManager, lpParentId As String, lpName As String, ByRef lpExistingFolderId As String) As Boolean
      Try
        If sender.FolderIDExists(lpParentId) = False Then
          Return False
        Else
          Dim lobjParentFolder As IFolder = sender.GetFolderInfoByID(lpParentId, 0)
          If lobjParentFolder.HasSubFolders = False Then
            Return False
          Else
            For Each lobjSubfolder As IFolder In lobjParentFolder.SubFolders
              If (String.Compare(lobjSubfolder.Name, lpName, True) = 0) _
                OrElse (String.Compare(lobjSubfolder.DisplayName, lpName, True) = 0) Then
                lpExistingFolderId = lobjSubfolder.Id
                Return True
              End If
            Next
            Return False
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Returns the folder specified
    ''' </summary>
    ''' <param name="lpFolderPath">The folder path to return</param>
    ''' <param name="lpMaxContentCount">Pass -1 to get all content, 0 to get no content, or any other number to specify a specific maximum</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function GetFolder(ByVal lpFolderPath As String, ByVal lpMaxContentCount As Long) As IFolder Implements IProvider.GetFolder

    ''' <summary>
    ''' Locates and returns a document based on the folder path.
    ''' </summary>
    ''' <param name="lpDocumentPath">The fully qualified folder path to the document, including the document name.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <exception cref="Exceptions.ProviderInterfaceNotImplementedException">
    ''' If the provider does not implement IDocumentExporter an ProviderInterfaceNotImplementedException will be thrown.
    ''' </exception>
    Public Function GetDocumentByPath(lpDocumentPath As String) As Document Implements IProvider.GetDocumentByPath
      Try

        If String.IsNullOrEmpty(lpDocumentPath) Then
          Throw New ArgumentNullException(NameOf(lpDocumentPath))
        End If

        Dim lobjExporter As IDocumentExporter = Me.GetInterface(Providers.ProviderClass.DocumentExporter)
        Dim lobjReturnDocument As Document = Nothing
        Dim lobjExportDocumentArgs As ExportDocumentEventArgs = Nothing
        Dim lblnExportDocumentSuccess As Boolean
        Dim lstrPathParts As String() = lpDocumentPath.Split(FolderDelimiter)
        Dim lstrDocumentName As String = Nothing
        If lstrPathParts.Length > 1 Then
          lstrDocumentName = lstrPathParts(lstrPathParts.Length - 1)
        Else
          Throw New Exceptions.InvalidPathException(lpDocumentPath)
        End If

        Dim lstrFolderPath As String = lpDocumentPath.Substring(0, (lpDocumentPath.Length - lstrDocumentName.Length - 1))

        Dim lobjFolder As IFolder = Me.GetFolder(lstrFolderPath, -1)

        If lobjFolder Is Nothing Then
          Throw New Exceptions.InvalidPathException(String.Format("No folder could be found for the path '{0}'.", lstrFolderPath), lstrFolderPath)
        End If

        If lobjFolder.ContentCount = 0 Then
          Throw New Exceptions.DocumentDoesNotExistException(lpDocumentPath, String.Format("The folder '{0}' has no contents.", lstrFolderPath))
        End If

        For Each lobjFolderContent As FolderContent In lobjFolder.Contents
          If String.Compare(lobjFolderContent.Name, lstrDocumentName, True) = 0 Then
            lobjExportDocumentArgs = New ExportDocumentEventArgs(lobjFolderContent)
            lblnExportDocumentSuccess = lobjExporter.ExportDocument(lobjExportDocumentArgs)
            If lblnExportDocumentSuccess = True Then
              Return lobjExportDocumentArgs.Document
            Else
              If Not String.IsNullOrEmpty(lobjExportDocumentArgs.ErrorMessage) Then
                Throw New Exceptions.DocumentException(lpDocumentPath, lobjExportDocumentArgs.ErrorMessage)
              Else
                Throw New Exceptions.DocumentException(lpDocumentPath, "Unable to get document by path.")
              End If
            End If
            Exit For
          End If
        Next

        Throw New Exceptions.DocumentDoesNotExistException(lpDocumentPath)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Methods"

    Public Overridable Sub Disconnect() Implements IProvider.Disconnect
      Try
        mblnIsConnected = False
        SetState(ProviderConnectionState.Disconnected)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Function FindInvalidProperties(ByVal lpDocument As Document,
                                                      ByVal lpPropertyScope As PropertyScope) As IProperties
      Try

        ' Make sure this provider supports classification
        If SupportsInterface(Providers.ProviderClass.Classification) = False Then
          ' This provider does not implement classification
          Throw New NotImplementedException(
            String.Format("Unable to check for invalid properties, the {0} provider does not support classification.", Me.Name))
        Else
          Dim lobjRequestedDocumentClass As DocumentClass
          Dim lobjInvalidProperties As InvalidProperties

          'lobjRequestedDocumentClass = CType(Me, IClassification).DocumentClasses(lpDocument.DocumentClass)
          lobjRequestedDocumentClass = CType(Me, IClassification).DocumentClass(lpDocument.DocumentClass)

          Dim lblnWarnForNonCompliance As Boolean = ProviderProperties(ENFORCE_CLASSIFICATION_COMPLIANCE).Value

          ' Check the properties based on the document class
          lobjInvalidProperties = lobjRequestedDocumentClass.FindInvalidProperties(
            lpDocument, PropertyScope.BothDocumentAndVersionProperties, lblnWarnForNonCompliance)


          '' Exclude all read only properties
          'For lintPropertyCounter As Integer = lobjInvalidProperties.Count - 1 To 0 Step -1
          '  If lobjRequestedDocumentClass.Properties.Contains(lobjInvalidProperties(lintPropertyCounter).Name) Then
          '    If lobjRequestedDocumentClass.Properties(lobjInvalidProperties(lintPropertyCounter).Name).Settability = ClassificationProperty.SettabilityEnum.READ_ONLY Then

          '    End If
          '  End If
          'Next

          ' Check to see if there are action properties in the list to be excluded
          'Dim lintPropertyCounter As Integer = lobjInvalidProperties.Count
          Dim lobjActionPropertyCandidate As IProperty = Nothing
          For lintPropertyCounter As Integer = lobjInvalidProperties.Count - 1 To 0 Step -1
            lobjActionPropertyCandidate = DirectCast(lobjInvalidProperties(lintPropertyCounter), IProperty)
            If Me.ActionProperties.Contains(lobjActionPropertyCandidate) Then
              lobjInvalidProperties.Remove(lobjInvalidProperties(lintPropertyCounter))
            End If
          Next

          ' See if we were requested to set system properties
          If lpDocument.GetFirstVersion.Properties.PropertyExists(ACTION_SET_SYSTEM_PROPERTIES, False) Then
            Dim lobjSetSystemProperties As Object = lpDocument.GetFirstVersion.Properties(ACTION_SET_SYSTEM_PROPERTIES).Value
            If lobjSetSystemProperties IsNot Nothing Then
              Try
                Me.ActionProperties(ACTION_SET_SYSTEM_PROPERTIES).Value = CBool(lobjSetSystemProperties)
              Catch ex As Exception
                Me.ActionProperties(ACTION_SET_SYSTEM_PROPERTIES).Value = False
              End Try
            End If
          End If

          Return lobjInvalidProperties

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Sub RemoveInvalidProperty(ByVal lpPropertyName As String,
                                        ByVal lpDocument As Core.Document,
                                        ByVal lpInvalidProperties As InvalidProperties,
                                        ByVal lpScope As InvalidProperty.InvalidPropertyScope)
      Try
        Dim lobjFoundProperty As ECMProperty = Nothing
        Dim lobjVersion As Version = Nothing


        Select Case lpScope
          Case InvalidProperty.InvalidPropertyScope.FirstVersion
            If lpDocument.FirstVersion.Properties.PropertyExists(lpPropertyName, False, CType(lobjFoundProperty, IProperty)) Then
              lpInvalidProperties.Add(lobjFoundProperty, InvalidProperty.InvalidPropertyScope.FirstVersion)
              lpDocument.FirstVersion.Properties.Delete(lpPropertyName)
              If LogInvalidPropertyRemovals Then
                ApplicationLogging.WriteLogEntry(
                  String.Format("Removed invalid property '{0}' from first version of document '{1}'",
                                lpPropertyName, lpDocument.Name), TraceEventType.Information, 4251)
              End If
            End If
          Case InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion
            For lintVersionCounter As Integer = lpDocument.Versions.Count - 1 To 1 Step -1
              lobjVersion = lpDocument.Versions(lintVersionCounter)
              If lobjVersion.Properties.PropertyExists(lpPropertyName, False, lobjFoundProperty) Then
                lpInvalidProperties.Add(lobjFoundProperty, InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion)
                lobjVersion.Properties.Delete(lpPropertyName)
                If LogInvalidPropertyRemovals Then
                  ApplicationLogging.WriteLogEntry(String.Format("Removed invalid property '{0}' from version {1} of document '{2}'",
                    lpPropertyName, lobjVersion.ID + 1, lpDocument.Name), TraceEventType.Information, 4252)
                End If
              End If
            Next
          Case InvalidProperty.InvalidPropertyScope.AllVersions
            For Each lobjDocumentVersion As Version In lpDocument.Versions
              If lobjDocumentVersion.Properties.PropertyExists(lpPropertyName, False, lobjFoundProperty) Then
                lpInvalidProperties.Add(lobjFoundProperty, InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion)
                lobjVersion.Properties.Delete(lpPropertyName)
                If LogInvalidPropertyRemovals Then
                  ApplicationLogging.WriteLogEntry(String.Format("Removed invalid property '{0}' from version {1} of document '{2}'",
                    lpPropertyName, lobjDocumentVersion.ID + 1, lpDocument.Name), TraceEventType.Information, 4253)
                End If
              End If
            Next
          Case InvalidProperty.InvalidPropertyScope.Document
            If lpDocument.Properties.PropertyExists(lpPropertyName, False, lobjFoundProperty) Then
              lpInvalidProperties.Add(lobjFoundProperty, InvalidProperty.InvalidPropertyScope.AllExceptFirstVersion)
              lpDocument.Properties.Delete(lpPropertyName)
              If LogInvalidPropertyRemovals Then
                ApplicationLogging.WriteLogEntry(
                  String.Format("Removed invalid property '{0}' from document '{1}'",
                                lpPropertyName, lpDocument.Name), TraceEventType.Information, 4250)
              End If
            End If
        End Select

        'lpInvalidProperties.Add(lobjFoundProperty)
        'lpDocument.DeleteProperty(PropertyScope.BothDocumentAndVersionProperties, lpPropertyName)
        'ApplicationLogging.WriteLogEntry( _
        '  String.Format("Removed invalid property '{0}' from document '{1}'", _
        '                lpPropertyName, lpDocument.Name), TraceEventType.Information, 4251)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates and returns a ContentSource connection string.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Generated connection string will not include a provider path.  
    ''' If a provider path is desired call GenerateConnectionString(True) instead.
    ''' </remarks>
    Public Overridable Function GenerateConnectionString() As String Implements IProvider.GenerateConnectionString

      Try
        Return GenerateConnectionString(ProviderProperties, False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::GenerateConnectionString", Me.GetType.Name))
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Creates and returns a ContentSource connection string.
    ''' </summary>
    ''' <param name="lpIncludeProviderPath">
    ''' Determines whether or not the generated 
    ''' connection string will include the provider path.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GenerateConnectionString(ByVal lpIncludeProviderPath As Boolean) As String Implements IProvider.GenerateConnectionString

      Try
        Return GenerateConnectionString(ProviderProperties, lpIncludeProviderPath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::GenerateConnectionString", Me.GetType.Name))
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Creates and returns a ContentSource connection string.
    ''' </summary>
    ''' <param name="lpProviderProperties">
    ''' A collection of provider properties detailing 
    ''' how to instantiate the ContentSource.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Generated connection string will not include a provider path.  
    ''' If a provider path is desired call GenerateConnectionString(lpProviderProperties, True) instead.
    ''' </remarks>
    Public Overridable Function GenerateConnectionString(ByVal lpProviderProperties As ProviderProperties) As String Implements IProvider.GenerateConnectionString

      Try
        Return GenerateConnectionString("", lpProviderProperties, False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::GenerateConnectionString(lpProviderProperties)", Me.GetType.Name))
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Creates and returns a ContentSource connection string.
    ''' </summary>
    ''' <param name="lpProviderProperties">
    ''' A collection of provider properties detailing 
    ''' how to instantiate the ContentSource.
    ''' </param>
    ''' <param name="lpIncludeProviderPath">
    ''' Determines whether or not the generated 
    ''' connection string will include the provider path.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GenerateConnectionString(ByVal lpProviderProperties As ProviderProperties,
                                                         ByVal lpIncludeProviderPath As Boolean) As String Implements IProvider.GenerateConnectionString

      Try
        Return GenerateConnectionString("", lpProviderProperties, lpIncludeProviderPath)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::GenerateConnectionString(lpProviderProperties)", Me.GetType.Name))
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Creates and returns a ContentSource connection string.
    ''' </summary>
    ''' <param name="lpContentSourceName">The name of the content source.</param>
    ''' <param name="lpProviderProperties">
    ''' A collection of provider properties detailing 
    ''' how to instantiate the ContentSource.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Generated connection string will not include a provider path.  
    ''' If a provider path is desired call 
    ''' GenerateConnectionString(lpContentSourceName, lpProviderProperties, True) instead.
    ''' </remarks>
    Public Overridable Function GenerateConnectionString(ByVal lpContentSourceName As String,
                                                         ByVal lpProviderProperties As ProviderProperties) As String Implements IProvider.GenerateConnectionString
      Try
        Return GenerateConnectionStringShared(lpContentSourceName, lpProviderProperties, Me.Name, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::GenerateConnectionString('{1}', lpProviderProperties)",
                                      Me.GetType.Name, lpContentSourceName))
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates and returns a ContentSource connection string.
    ''' </summary>
    ''' <param name="lpContentSourceName">The name of the content source.</param>
    ''' <param name="lpProviderProperties">
    ''' A collection of provider properties detailing 
    ''' how to instantiate the ContentSource.
    ''' </param>
    ''' <param name="lpIncludeProviderPath">
    ''' Determines whether or not the generated 
    ''' connection string will include the provider path.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GenerateConnectionString(ByVal lpContentSourceName As String,
                                                         ByVal lpProviderProperties As ProviderProperties,
                                                         ByVal lpIncludeProviderPath As Boolean) As String Implements IProvider.GenerateConnectionString
      Try
        If lpIncludeProviderPath = True Then
          Return GenerateConnectionStringShared(lpContentSourceName, lpProviderProperties, Me.Name, Me.ProviderPath)
        Else
          Return GenerateConnectionStringShared(lpContentSourceName, lpProviderProperties, Me.Name, String.Empty)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::GenerateConnectionString('{1}', lpProviderProperties)",
                                      Me.GetType.Name, lpContentSourceName))
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates and returns a ContentSource connection string.
    ''' </summary>
    ''' <param name="lpContentSourceName">The name of the content source.</param>
    ''' <param name="lpProviderProperties">
    ''' A collection of provider properties detailing 
    ''' how to instantiate the ContentSource.
    ''' </param>
    ''' <param name="lpProviderName">The name of the provider.</param>
    ''' <param name="lpProviderPath">
    ''' The fully qualified path to the provider file.  
    ''' If the supplied value is null or an empty string, 
    ''' then the provider path will not be included as 
    ''' part of the generated connection string.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GenerateConnectionStringShared(ByVal lpContentSourceName As String,
                                                          ByVal lpProviderProperties As ProviderProperties,
                                                          ByVal lpProviderName As String,
                                                          ByVal lpProviderPath As String) As String
      Try

        ' Set the name of the content source
        Dim lstrConnectionString As String = "Name=" & lpContentSourceName

        ' Set the name of the provider
        lstrConnectionString &= ";" & "Provider=" & lpProviderName

        ' Set all the remaining provider properties
        For Each lobjProviderProperty As ProviderProperty In lpProviderProperties

          'If lobjProviderProperty.PropertyName.ToUpper = "PROVIDERPATH" Then
          If String.Equals(lobjProviderProperty.PropertyName, "ProviderPath", StringComparison.InvariantCultureIgnoreCase) Then
            ' If a provider path was not specified then we will not include 
            ' it even if it was included as one of the provider properties.
            If lpProviderPath Is Nothing OrElse lpProviderPath.Length = 0 Then
              Continue For
            End If
          End If

          'If lobjProviderProperty.PropertyName.Equals(PWD, StringComparison.InvariantCultureIgnoreCase) = True Then
          If lobjProviderProperty.PropertyName.ToLower.Contains(PWD.ToLower) = True Then
            ' Encrypt the password before putting it into the connection string
            'Dim lstrEncryptedPassword As String = Core.Password.EscapedEncrypt(lobjProviderProperty.PropertyValue)
            '  Changed by EFB on 4/27/08
            '  Change from Escaped Encrypted value to HEX representation of encrypted value.
            Dim lstrEncryptedPassword As String = Core.Password.Encrypt(lobjProviderProperty.PropertyValue).Hex
            lstrConnectionString &= ";" & lobjProviderProperty.PropertyName & "=" & lstrEncryptedPassword
          ElseIf lobjProviderProperty.PropertyName.ToUpper.Contains("CONNECTIONSTRING") = True Then
            ' If this is a nested connection string then enclose it in curly braces
            lstrConnectionString &= ";" & lobjProviderProperty.PropertyName & "={" & lobjProviderProperty.PropertyValue & "}"
          ElseIf String.Equals(lobjProviderProperty.PropertyName, EXPORT_PATH, StringComparison.InvariantCultureIgnoreCase) OrElse
            String.Equals(lobjProviderProperty.PropertyName, IMPORT_PATH, StringComparison.InvariantCultureIgnoreCase) Then
            If lobjProviderProperty.PropertyValue IsNot Nothing AndAlso TypeOf lobjProviderProperty.PropertyValue Is String Then
              lobjProviderProperty.PropertyValue = lobjProviderProperty.PropertyValue.ToString.Replace(
                FileHelper.Instance.CtsDocsPath, CTS_DOCS_PATH_REPLACEMENT)
            End If
            lstrConnectionString &= ";" & lobjProviderProperty.PropertyName & "=" & lobjProviderProperty.PropertyValue
          Else
            If lstrConnectionString.Contains(lobjProviderProperty.PropertyName & "=") = False Then
              ' We don't want to duplicate entries
              lstrConnectionString &= ";" & lobjProviderProperty.PropertyName & "=" & lobjProviderProperty.PropertyValue
            End If
          End If

        Next

#If NET8_0_OR_GREATER Then
        If lstrConnectionString.Contains("PROVIDERPATH", StringComparison.CurrentCultureIgnoreCase) = False Then
          If Not String.IsNullOrEmpty(lpProviderPath) Then
            lstrConnectionString &= ";" & "ProviderPath" & "=" & lpProviderPath
          End If
        End If
#Else
        If lstrConnectionString.ToUpper.Contains("PROVIDERPATH") = False Then
          If Not String.IsNullOrEmpty(lpProviderPath) Then
            lstrConnectionString &= ";" & "ProviderPath" & "=" & lpProviderPath
          End If
        End If
#End If

        Return lstrConnectionString
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format($"CProvider::GenerateConnectionString('{1}', lpProviderProperties)",
                                      lpContentSourceName))
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetInterfaceDictionary() As Dictionary(Of ProviderClass, Boolean)
      Try
        Dim lobjDictionary As New Dictionary(Of ProviderClass, Boolean)

        Dim lobjProviderClasses() As ProviderClass = [Enum].GetValues(GetType(ProviderClass))

        For Each lenuProviderClass As ProviderClass In lobjProviderClasses
          lobjDictionary.Add(lenuProviderClass, SupportsInterface(lenuProviderClass))
        Next

        Return lobjDictionary

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Used to determine whether or not the current provider instance supports the requested interface.
    ''' </summary>
    ''' <param name="lpInterface">The desired interface as defined in the ProviderClass enumeration.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SupportsInterface(ByVal lpInterface As ProviderClass) As Boolean Implements IProvider.SupportsInterface

      ' What interfaces does this provider support?
      Dim lobjAssembly As System.Reflection.Assembly
      Dim lobjProviderCandidate As Type

      Try

        ' Validate the Provider
        Try
          lobjAssembly = Reflection.Assembly.GetAssembly(Me.GetType)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::SupportsInterface('{1}')_GetAssembly",
                                        Me.GetType.Name, lpInterface))
          ' If we are unable to load the assembly then it is not valid.
          Return False
        End Try

        Dim lstrInterfaceName As String
        Select Case lpInterface
          ' Explorer
          Case ProviderClass.Explorer
            lstrInterfaceName = "IExplorer"

            ' Exporter
          Case ProviderClass.DocumentExporter
            lstrInterfaceName = "IDocumentExporter"

            ' Importer
          Case ProviderClass.DocumentImporter
            lstrInterfaceName = "IDocumentImporter"

            ' Classification
          Case ProviderClass.Classification
            lstrInterfaceName = "IClassification"

            ' Basic Content Services
          Case ProviderClass.BasicContentServices
            lstrInterfaceName = "IBasicContentServicesProvider"

            ' Records Manager
          Case ProviderClass.RecordsManager
            lstrInterfaceName = "IRecordsManager"

          Case ProviderClass.RepositoryDiscovery
            lstrInterfaceName = "IRepositoryDiscovery"

            ' ProviderImages
          Case ProviderClass.ProviderImages
            lstrInterfaceName = "IProviderImages"

            ' Copy
          Case ProviderClass.Copy
            lstrInterfaceName = "ICopy"

            ' Delete
          Case ProviderClass.Delete
            lstrInterfaceName = "IDelete"

            ' Rename
          Case ProviderClass.Rename
            lstrInterfaceName = "IRename"

            ' Version
          Case ProviderClass.Version
            lstrInterfaceName = "IVersion"

            ' File
          Case ProviderClass.File
            lstrInterfaceName = "IFile"

            ' FolderManager
          Case ProviderClass.FolderManager
            lstrInterfaceName = "IFolderManager"

            ' FolderClassification
          Case ProviderClass.FolderClassification
            lstrInterfaceName = "IFolderClassification"

            ' SecurityClassification
          Case ProviderClass.SecurityClassification
            lstrInterfaceName = "ISecurityClassification"

            ' Create
          Case ProviderClass.Create
            lstrInterfaceName = "ICreate"

            'UpdateProperties
          Case ProviderClass.UpdateProperties
            lstrInterfaceName = "IUpdateProperties"

            'UpdatePermissions
          Case ProviderClass.UpdatePermissions
            lstrInterfaceName = "IUpdatePermissions"

          Case ProviderClass.ChoiceListExporter
            lstrInterfaceName = "IChoiceListExporter"

          Case ProviderClass.ChoiceListImporter
            lstrInterfaceName = "IChoiceListImporter"

          Case ProviderClass.SQLPassThroughSearch
            lstrInterfaceName = "ISQLPassThroughSearch"

          Case ProviderClass.LinkManager
            lstrInterfaceName = "ILinkManager"

          Case ProviderClass.AnnotationExporter
            lstrInterfaceName = "IAnnotationExporter"

          Case ProviderClass.AnnotationImporter
            lstrInterfaceName = "IAnnotationImporter"

          Case ProviderClass.DocumentClassImporter
            lstrInterfaceName = "IDocumentClassImporter"

          Case ProviderClass.FolderExporter
            lstrInterfaceName = "IFolderExporter"

          Case ProviderClass.FolderImporter
            lstrInterfaceName = "IFolderImporter"

          Case ProviderClass.CustomObjectExporter
            lstrInterfaceName = "ICustomObjectExporter"

          Case ProviderClass.CustomObjectImporter
            lstrInterfaceName = "ICustomObjectImporter"

          Case ProviderClass.CustomObjectClassification
            lstrInterfaceName = "ICustomObjectClassification"

          Case Else
            Throw New ArgumentOutOfRangeException(NameOf(lpInterface), String.Format("Provider class {0} is not yet covered in SupportsInterface.", lpInterface.ToString()))

        End Select

        'If Information.Interfaces.Contains(lstrInterfaceName) Then
        '  Return True
        'Else
        '  Return False
        'End If

        For Each lobjType As Type In lobjAssembly.GetTypes
          ' Check for interface
          lobjProviderCandidate = lobjType.GetInterface(lstrInterfaceName)
          If lobjProviderCandidate IsNot Nothing Then
            Return True
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::SupportsInterface('{1}')",
                                      Me.GetType.Name, lpInterface))
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <Added by: Ernie at: 9/28/2011-7:55:20 AM on machine: ERNIE-M4400>

    ''' <summary>
    ''' Forces the initialization of the classification related properties so 
    ''' that they may be initialized in advance instead of upon first request.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub InitializeClassificationProperties() Implements IProvider.InitializeClassificationProperties
      Try

        If mblnClassificationPropertiesInitialized = False AndAlso TypeOf Me Is IClassification Then
          ' I implement IClassification and I have not yet initialized the associated properties

          If Me.IsInitialized = False Then
            InitializeProvider(Me.ContentSource)
          End If

          If Me.IsConnected = False Then
            Connect()
          End If

          ' Create a reference to my classification interface
          Dim lobjClassifier As IClassification = DirectCast(Me, IClassification)

          ' Force the retrieval of document classes
          Dim lobjDocumentClasses As DocumentClasses = lobjClassifier.DocumentClasses

          ' Force the retrieval of all content properties
          Dim lobjContentProperties As ClassificationProperties = lobjClassifier.ContentProperties

          mblnClassificationPropertiesInitialized = True

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub
    ' </Added by: Ernie at: 9/28/2011-7:55:20 AM on machine: ERNIE-M4400>

    Public Overridable Sub InitializeProvider(ByVal lpContentSource As ContentSource) Implements IProvider.InitializeProvider

      Try
        'ApplicationLogging.WriteLogEntry(String.Format("Enter CProvider::{0}::InitializeProvider", Me.GetType.Name), TraceEventType.Verbose)

        If lpContentSource Is Nothing Then
          Throw New ArgumentNullException(NameOf(lpContentSource), "Unable to initialize provider, the content source parameter is null.")
        End If

        For Each lobjProperty As ProviderProperty In lpContentSource.Properties
          Select Case lobjProperty.PropertyName
            Case "Provider", "Name", "ProviderPath", IMPORT_PATH
              ' Ignore this one since it only applies to Content Sources and not providers

            Case ACTION_SET_SYSTEM_PROPERTIES
              ActionProperties(ACTION_SET_SYSTEM_PROPERTIES).Value = lobjProperty.PropertyValue

            Case Else
              Try
                'If lobjProperty.PropertyName.Equals(PWD, StringComparison.InvariantCultureIgnoreCase) = True Then
                If lobjProperty.PropertyName.ToLower.Contains(PWD.ToLower) = True Then
                  Try
                    'Dim lstrDecryptedPassword As String = Core.Password.EscapedDecrypt(lobjProperty.PropertyValue)
                    '  Changed by EFB on 4/27/08
                    '  Change from Escaped Encrypted value to HEX representation of encrypted value.
                    Dim lstrDecryptedPassword As String = Core.Password.DecryptStringFromHex(lobjProperty.PropertyValue)

                    Dim lobjProviderProperty As ProviderProperty = ProviderProperties(lobjProperty.PropertyName)
                    If lobjProviderProperty IsNot Nothing Then
                      lobjProviderProperty.PropertyValue = lstrDecryptedPassword
                    Else
                      lobjProperty.PropertyValue = lstrDecryptedPassword
                      ProviderProperties.Add(lobjProperty)
                      'ApplicationLogging.WriteLogEntry(String.Format( _
                      '  "Unable to initialize provider property '{0}', the property could not be found for provider '{1}'.", _
                      '  lobjProperty.PropertyName, Me.Name), TraceEventType.Warning, 5693)
                    End If

                    If lobjProperty.PropertyName.Equals(PWD, StringComparison.InvariantCultureIgnoreCase) = True Then
                      Me.Password = lstrDecryptedPassword
                    End If

                  Catch ex As System.Security.Cryptography.CryptographicUnexpectedOperationException
                    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
                    ProviderProperties(lobjProperty.PropertyName).PropertyValue = lobjProperty.PropertyValue
                    If lobjProperty.PropertyName.Equals(PWD, StringComparison.InvariantCultureIgnoreCase) = True Then
                      Me.Password = lobjProperty.PropertyValue
                    End If
                  Catch ex As Exception
                    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
                    ProviderProperties(lobjProperty.PropertyName).PropertyValue = lobjProperty.PropertyValue
                    If lobjProperty.PropertyName.Equals(PWD, StringComparison.InvariantCultureIgnoreCase) = True Then
                      Me.Password = lobjProperty.PropertyValue
                    End If
                  End Try
                Else
                  Dim lobjProviderProperty As ProviderProperty = ProviderProperties(lobjProperty.PropertyName)
                  If lobjProperty.PropertyName.Equals("username", StringComparison.CurrentCultureIgnoreCase) Then
                    Try
                      Me.UserName = lobjProperty.PropertyValue
                    Catch ex As Exception
                      ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod())
                      ' Skip it and move on
                    End Try
                  End If
                  If lobjProviderProperty IsNot Nothing Then
                    lobjProviderProperty.PropertyValue = lobjProperty.PropertyValue
                  Else
                    ProviderProperties.Add(lobjProperty)
                    'ApplicationLogging.WriteLogEntry(String.Format( _
                    '  "Unable to initialize provider property '{0}', the property could not be found for provider '{1}'.", _
                    '  lobjProperty.PropertyName, Me.Name), TraceEventType.Warning, 5693)
                  End If
                  'ProviderProperties(lobjProperty.PropertyName).PropertyValue = lobjProperty.PropertyValue
                End If
              Catch ex As Exception
                ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::InitializeProvider", Me.GetType.Name))
                ' We could not find this one, skip it for now.
                'Re-throw the exception to the caller
                Throw
              End Try
          End Select
        Next

        'If SupportsInterface(Providers.ProviderClass.DocumentExporter) Then
        '  ' Attempt to set the ExportPath
        '  Try
        '    ExportPath = lpContentSource.ExportPath
        '  Catch ex As Exception
        '    ApplicationLogging.LogException(ex, String.Format("Error setting the ExportPath in CProvider::{0}::InitializeProvider", Me.GetType.Name))
        '    ' Unable to set the ExportPath from the data source
        '  End Try
        'End If

        If SupportsInterface(Providers.ProviderClass.DocumentImporter) Then
          If SupportsInterface(ProviderClass.Classification) Then
            Dim lobjDisableInvalidPropertyChecksBehaviorProperty As ProviderProperty = ProviderProperties.Item(ENFORCE_CLASSIFICATION_COMPLIANCE)
            If lobjDisableInvalidPropertyChecksBehaviorProperty IsNot Nothing Then
              Try
                mblnEnforceClassificationCompliance = Boolean.Parse(lobjDisableInvalidPropertyChecksBehaviorProperty.PropertyValue)
              Catch ex As Exception
                ApplicationLogging.LogException(ex,
                  String.Format("Error setting the {0} property to '{1}' in CProvider.{2}.InitializeProvider",
                                ENFORCE_CLASSIFICATION_COMPLIANCE, lobjDisableInvalidPropertyChecksBehaviorProperty.PropertyValue, Me.GetType.Name))
              End Try
            End If
            Dim lobjInvalidPropertyLoggingBehaviorProperty As ProviderProperty = ProviderProperties.Item(LOG_INVALID_PROPERTY_REMOVALS)
            If lobjInvalidPropertyLoggingBehaviorProperty IsNot Nothing Then
              Try
                mblnLogInvalidPropertyRemovals = Boolean.Parse(lobjInvalidPropertyLoggingBehaviorProperty.PropertyValue)
              Catch ex As Exception
                ApplicationLogging.LogException(ex,
                  String.Format("Error setting the {0} property to '{1}' in CProvider.{2}.InitializeProvider",
                                LOG_INVALID_PROPERTY_REMOVALS, lobjInvalidPropertyLoggingBehaviorProperty.PropertyValue, Me.GetType.Name))
              End Try
            Else
              mblnEnforceClassificationCompliance = False
              mblnLogInvalidPropertyRemovals = False
            End If
          End If
          ' <Removed by: Ernie at: 9/29/2014-11:17:40 AM on machine: ERNIE-THINK>
          '           ' Attempt to set the ImportPath
          '           Try
          '             ImportPath = lpContentSource.ImportPath
          '           Catch ex As Exception
          '             ApplicationLogging.LogException(ex, String.Format("Error setting the ImportPath in CProvider::{0}::InitializeProvider", Me.GetType.Name))
          '             ' Unable to set the ImportPath from the data source
          '           End Try
          ' </Removed by: Ernie at: 9/29/2014-11:17:40 AM on machine: ERNIE-THINK>
        End If

        If Me.ContentSource Is Nothing Then
          Me.ContentSource = lpContentSource
        End If

        mblnIsInitialized = True

      Catch ex As Exception
        mblnIsInitialized = False
        ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::InitializeProvider", Me.GetType.Name))
        '  Re-throw the exception to the caller
        Throw
      Finally
        'ApplicationLogging.WriteLogEntry(String.Format("Exit CProvider::{0}::InitializeProvider", Me.GetType.Name), TraceEventType.Verbose)
      End Try

    End Sub

    ''' <summary>
    ''' Updates the system properties
    ''' </summary>
    ''' <param name="Args">UpdateSystemProperties Argument</param>
    ''' <param name="lpDelegate">The address of the delegate method that will actually perform the updates</param>
    ''' <remarks></remarks>
    Public Overridable Sub UpdateSystemProperties(ByVal Args As UpdateSystemPropertiesArgs,
                                                  ByVal lpDelegate As UpdateSystemPropertiesCallback)
      Try
        lpDelegate.Invoke(Args)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Returns a fresh search
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride Function CreateSearch() As ISearch Implements IProvider.CreateSearch

    'Public Overrides Function ToString() As String
    '  Try
    '    If Me.Name IsNot Nothing AndAlso Me.Name.Length > 0 Then
    '      Return Me.Name
    '    Else
    '      Return Me.GetType.Name
    '    End If

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::ToString", Me.GetType.Name))
    '    Return ""
    '  End Try
    'End Function

#End Region

#Region "Protected Methods"

    Delegate Sub EventForwarderCallback(ByVal sender As Object, ByVal e As EventArgs)
    Delegate Function ExportDocumentCallback(ByVal e As ExportDocumentEventArgs) As Boolean

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        If Me.Name IsNot Nothing AndAlso Me.Name.Length > 0 Then
          Return Me.Name
        Else
          Return Me.GetType.Name
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::ToString", Me.GetType.Name))
        Return String.Empty
      End Try
    End Function

    ' <Removed by: Ernie at: 9/29/2014-1:52:46 PM on machine: ERNIE-THINK>
    '     Protected Sub ExportFolder(ByVal sender As IDocumentExporter, ByVal Args As ExportFolderEventArgs)
    ' 
    '       Try
    ' 
    '         Dim lintContentCount As Long = 0
    '         Dim lintContentCounter As Integer = 0
    '         Dim lobjECMDocument As Document = Nothing
    '         Dim lenuRecursionLevel As RecursionLevel = Args.RecursionLevel
    '         Dim lblnExportSuccess As Boolean
    '         Dim lobjExportDocumentArgs As ExportDocumentEventArgs
    ' 
    '         Dim lintSuccessCounter As Integer
    '         Dim lblnIsEvaluation As Boolean = LicenseRegister.IsEvaluation
    ' 
    '         Dim lobjExporter As IDocumentExporter = sender
    ' 
    '         Select Case Args.FolderSource
    '           Case ExportFolderEventArgs.FolderSourceType.FolderReference
    '             lintContentCount = Args.Folder.Contents.Count
    '             Dim lobjFolderContents As FolderContents = Args.Folder.Contents
    '             Dim lobjFolder As IFolder = Args.Folder
    ' 
    '             If Args.WorkInBackground = True Then
    ' 
    ' 
    '               ' Export each document in the Folder
    '               For Each lobjContentItem As FolderContent In lobjFolderContents
    '                 ' Check for cancellation.
    '                 If Args.Worker.CancellationPending Then
    '                   Args.DoWorkEventArgs.Cancel = True
    '                   Exit Sub
    '                 End If
    '                 'Dim lobjExportDocumentArgs As New ExportDocumentArgs(lobjContentItem.ID, lobjECMDocument, Args.Worker)
    '                 'Dim lobjExportDocumentArgs As New ExportDocumentEventArgs(lobjContentItem, lobjECMDocument, Args.Worker, Args.Transformation)
    '                 lobjExportDocumentArgs = New ExportDocumentEventArgs(lobjContentItem, lobjECMDocument, Args.Worker, Args.Transformation)
    ' 
    '                 ' Make sure we write the cdf to disk.
    '                 lobjExportDocumentArgs.GenerateCDF = True
    ' 
    '                 Try
    '                   lblnExportSuccess = lobjExporter.ExportDocument(lobjExportDocumentArgs)
    ' 
    '                   If lblnExportSuccess = True Then
    '                     ' Write a success entry in the log
    '                     ApplicationLogging.WriteLogEntry( _
    '                       String.Format("Successfully exported document '{0}'.", _
    '                                     lobjExportDocumentArgs.Id), _
    '                       TraceEventType.Information, EXPORT_DOCUMENT_SUCCESS)
    '                     TransactionCounter.Instance.Increment()
    '                     lintSuccessCounter += 1
    '                     If lblnIsEvaluation AndAlso lintSuccessCounter = 100 Then
    '                       Exit For
    '                     End If
    '                   End If
    ' 
    '                 Catch ex As Exception
    '                   ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '                   ' Write an error file to the Export Path
    '                   'WriteExportError(lobjContentItem.ID, Ecmg.Utilities.Helper.FormatCallStack(ex))
    '                   Dim lobjWriteErrorArgs As New WriteErrorEventArgs(Helper.FormatCallStack(ex))
    '                   lobjWriteErrorArgs.Id = lobjContentItem.ID
    '                   ' lpCallBack.Invoke(sender, lobjWriteErrorArgs)
    '                   sender.OnDocumentExportError(New DocumentExportErrorEventArgs(lobjContentItem.ID, Helper.FormatCallStack(ex), ex))
    '                 End Try
    '                 lintContentCounter += 1
    ' 
    '                 'RKS 4/3/2008 - not sure if this is causing the duplicate calls but
    '                 'going to comment out for now and see if it works
    '                 Dim lobjFolderDocumentExportedEventArgs As New FolderDocumentExportedEventArgs(lobjExportDocumentArgs.Document, Now, lintContentCounter, lintContentCount, Args.Worker)
    '                 sender.OnFolderDocumentExported(lobjFolderDocumentExportedEventArgs)
    '               Next
    ' 
    '               ' Depending on the recursion level export the child folders
    '               If lenuRecursionLevel <> RecursionLevel.ecmThisLevelOnly AndAlso lobjFolder.HasSubFolders = True Then
    '                 For Each lobjIdmSubFolder As IFolder In lobjFolder.SubFolders
    '                   If lenuRecursionLevel <> RecursionLevel.ecmAllChildren Then
    '                     ' Recurse only through the immediate child folders
    '                     lobjExporter.ExportFolder(New ExportFolderEventArgs(lobjIdmSubFolder, RecursionLevel.ecmThisLevelOnly, Args.Worker))
    '                   Else
    '                     ' Recurse through all levels
    '                     lobjExporter.ExportFolder(New ExportFolderEventArgs(lobjIdmSubFolder, RecursionLevel.ecmAllChildren, Args.Worker))
    '                   End If
    '                 Next
    '               End If
    ' 
    '               ' Notify calling program
    '               'RaiseEvent FolderExported(lobjFolder.Path, Now)
    '               'lpCallBack.Invoke(sender, New FolderExportedEventArgs(lobjFolder, Now))
    '               ApplicationLogging.WriteLogEntry(String.Format("'{0}' folder exported successfully to '{1}'!", Args.Folder.Path, sender.ExportPath), TraceEventType.Information)
    '               sender.OnFolderExported(New FolderExportedEventArgs(lobjFolder, sender.ExportPath, Now))
    '               Args.DoWorkEventArgs.Result = "Folder Exported: " & lobjFolder.Path & vbCrLf & Now
    ' 
    '             Else 'Args.WorkInBackground = False
    '               'Dim lintContentCount As Long = 0
    '               'Dim lintContentCounter As Integer = 0
    '               'Dim lobjECMDocument As Document = Nothing
    ' 
    '               ' Export each document in the Folder
    '               For Each lobjContentItem As FolderContent In lobjFolderContents
    '                 'ExportDocument(lobjContentItem.ID, lobjECMDocument)
    '                 lobjExportDocumentArgs = New ExportDocumentEventArgs(lobjContentItem.ID, lobjECMDocument)
    '                 lobjExportDocumentArgs.GenerateCDF = True
    '                 lblnExportSuccess = lobjExporter.ExportDocument(lobjExportDocumentArgs)
    ' 
    '                 If lblnExportSuccess = True Then
    '                   ' Write a success entry in the log
    '                   ApplicationLogging.WriteLogEntry( _
    '                     String.Format("Successfully exported document '{0}'.", _
    '                                   lobjExportDocumentArgs.Id), _
    '                     TraceEventType.Information, EXPORT_DOCUMENT_SUCCESS)
    '                   TransactionCounter.Instance.Increment()
    '                   lintSuccessCounter += 1
    '                   If lblnIsEvaluation AndAlso lintSuccessCounter = 100 Then
    '                     Exit For
    '                   End If
    '                 End If
    ' 
    '                 lintContentCounter += 1
    '                 'RaiseEvent FolderDocumentExported(lobjECMDocument, Now, lintContentCounter, lintContentCount, Nothing, Nothing)
    '                 'lpCallBack.Invoke(sender, New FolderDocumentExportedEventArgs(lobjECMDocument, Now, lintContentCounter, lintContentCount, Nothing))
    '                 sender.OnFolderDocumentExported(New FolderDocumentExportedEventArgs(lobjExportDocumentArgs.Document, Now, lintContentCounter, lintContentCount, Nothing))
    '               Next
    ' 
    '               ' Depending on the recursion level export the child folders
    '               If lenuRecursionLevel <> RecursionLevel.ecmThisLevelOnly AndAlso lobjFolder.HasSubFolders = True Then
    '                 For Each lobjIdmSubFolder As IFolder In lobjFolder.SubFolders
    '                   If lenuRecursionLevel <> RecursionLevel.ecmAllChildren Then
    '                     ' Recurse only through the immediate child folders
    '                     lobjExporter.ExportFolder(New ExportFolderEventArgs(lobjIdmSubFolder, RecursionLevel.ecmThisLevelOnly))
    '                   Else
    '                     ' Recurse through all levels
    '                     lobjExporter.ExportFolder(New ExportFolderEventArgs(lobjIdmSubFolder, RecursionLevel.ecmAllChildren))
    '                   End If
    '                 Next
    '               End If
    ' 
    '               ' Notify calling program
    '               'RaiseEvent FolderExported(lobjFolder.Path, Now)
    '               'lpCallBack.Invoke(sender, New FolderExportedEventArgs(lobjFolder, Now))
    '               ApplicationLogging.WriteLogEntry(String.Format("Folder Exported: {0}", lobjFolder.Path), TraceEventType.Information)
    '               sender.OnFolderExported(New FolderExportedEventArgs(lobjFolder, sender.ExportPath, Now))
    '               Args.DoWorkEventArgs.Result = "Folder Exported: " & lobjFolder.Path & vbCrLf & Now
    ' 
    '             End If
    ' 
    '           Case ExportFolderEventArgs.FolderSourceType.FolderPath
    ' 
    '             Dim lobjFolder As IFolder
    '             '          Try
    '             'lobjFolder = CType(sender, IExplorer).RootFolder.SubFolders(Args.FolderPath)
    '             lobjFolder = CType(sender, CProvider).GetFolder(Args.FolderPath, -1)
    '             'lobjFolder = Me.GetFolder(Args.FolderPath)
    '             'Catch ex As Exception
    '             '  '  Beep()
    '             'End Try
    ' 
    ' 
    '             If Args.WorkInBackground = True Then
    ' 
    '               Try
    '                 Dim lobjExportFolderArgs As New ExportFolderEventArgs(lobjFolder, Args.RecursionLevel, Args.Worker)
    '                 lobjExporter.ExportFolder(lobjExportFolderArgs)
    '               Catch ex As Exception
    '                 ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '                 Throw New ApplicationException("Unable to export folder by path: " & Args.FolderPath, ex)
    '               End Try
    ' 
    '             Else
    ' 
    '               Try
    '                 lobjExporter.ExportFolder(New ExportFolderEventArgs(lobjFolder, Args.RecursionLevel))
    '               Catch ex As Exception
    '                 ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '                 Throw New ApplicationException("Unable to export folder by path: " & Args.FolderPath, ex)
    '               End Try
    ' 
    '             End If
    ' 
    '         End Select
    '       Catch ex As Exception
    '         ApplicationLogging.LogException(ex, String.Format("CProvider::{0}::ExportFolder", Me.GetType.Name))
    '         ' Pass it on to the caller
    '         Throw
    '       End Try
    ' 
    '     End Sub
    ' </Removed by: Ernie at: 9/29/2014-1:52:46 PM on machine: ERNIE-THINK>

    ''' <summary>
    ''' To be called upon the completion of ExportDocument to centralize the serialization of the document if required.
    ''' </summary>
    ''' <param name="Args">The same argument passed into ExportDocument</param>
    ''' <remarks></remarks>
    Protected Function ExportDocumentComplete(ByVal sender As IDocumentExporter, ByVal Args As ExportDocumentEventArgs) As Boolean
      Try

        Dim lobjDocument As Document = Args.Document
        Dim lstrDocumentPath As String = String.Empty
        Dim lstrFileName As String = String.Empty
        Dim lstrArchivePath As String = String.Empty
        Dim lstrExportPath As String = ResolvedExportPath

        If (lobjDocument.Versions.Count > 0) Then
          ' Sort the document properties
          lobjDocument.Properties.Sort()

          ' Sort the versions
          lobjDocument.Versions.Sort()

        Else
          Args.Document = lobjDocument
          Args.ErrorMessage = "No versions available for document " & lobjDocument.ID
          Dim lobjDocEx As New Exceptions.DocumentException(Args.Id, Args.ErrorMessage)
          sender.OnDocumentExportError(New DocumentExportErrorEventArgs(Args, lobjDocEx))
          Return False
        End If

        ' If a transformation is present, then transform the document
        If (Args.Transformation IsNot Nothing) Then
          lobjDocument = lobjDocument.Transform(Args.Transformation)
        End If

        If Args.GetContent = True Then
          If (Args.GenerateCDF = True) OrElse (Args.Archive = True) Then
            If ExportAsArchive = True Then
              lstrDocumentPath = lstrExportPath
            Else
              lstrDocumentPath = String.Format("{0}{1}", lstrExportPath, lobjDocument.ID)
            End If

            If IO.Directory.Exists(lstrDocumentPath) = False Then
              Try
                IO.Directory.CreateDirectory(lstrDocumentPath)
              Catch ex As Exception
                ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
                ApplicationLogging.WriteLogEntry("Unable to write document to file, filed to create destination directory", TraceEventType.Error, 64723)
                Exit Function
              End Try
            End If

          End If

          If Args.Archive = True Then
            ' Archive the document
            lstrFileName = Helper.CleanFile(String.Format("{0}.{1}", lobjDocument.ID, Document.CONTENT_PACKAGE_FILE_EXTENSION), "^")
            lstrArchivePath = String.Format("{0}{1}", lstrDocumentPath, lstrFileName)
            lobjDocument.Archive(lstrArchivePath, True)
          End If

          If Args.GenerateCDF = True AndAlso Args.Archive = False AndAlso ExportAsArchive = True Then
            ' Archive the document
            lstrFileName = Helper.CleanFile(String.Format("{0}.{1}", lobjDocument.ID, Document.CONTENT_PACKAGE_FILE_EXTENSION), "^")
            lstrArchivePath = String.Format("{0}{1}", lstrDocumentPath, lstrFileName)
            lobjDocument.Archive(lstrArchivePath, True)
          ElseIf Args.GenerateCDF = True Then
            ' Serialize the document
            lstrFileName = Helper.CleanFile(String.Format("{0}.{1}", lobjDocument.ID, Document.CONTENT_DEFINITION_FILE_EXTENSION), "^")
            lobjDocument.Save(String.Format("{0}\{1}", lstrDocumentPath, lstrFileName))
          End If

          If lstrArchivePath.Length > 0 Then
            lobjDocument.SerializationPath = lstrArchivePath
          End If

        End If

        If String.IsNullOrEmpty(Args.ErrorMessage) Then
          sender.OnDocumentExported(New DocumentExportedEventArgs(lobjDocument, Now, Args.ErrorMessage, Args.Worker))
          Return True
        Else
          sender.OnDocumentExportMessage(New WriteMessageArgs(Args.ErrorMessage))
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' ''' <summary>
    ' ''' To be called upon the completion of ExportDocument to centralize the serialization of the folder if required.
    ' ''' </summary>
    ' ''' <param name="Args">The same argument passed into ExportDocument</param>
    ' ''' <remarks></remarks>
    'Protected Function ExportFolderComplete(ByVal sender As IDocumentExporter, ByVal Args As ExportFolderEventArgs) As Boolean
    '  Try

    '    Dim lobjFolder As Folder = Args.Folder
    '    Dim lstrFolderPath As String = String.Empty
    '    Dim lstrFileName As String = String.Empty
    '    Dim lstrArchivePath As String = String.Empty
    '    Dim lstrExportPath As String = ResolvedExportPath

    '    lobjFolder.Properties.Sort()

    '    '' If a transformation is present, then transform the document
    '    'If (Args.Transformation IsNot Nothing) Then
    '    '  lobjDocument = lobjDocument.Transform(Args.Transformation)
    '    'End If

    '    If Args.GetContent = True Then
    '      If (Args.GenerateCDF = True) OrElse (Args.Archive = True) Then
    '        If ExportAsArchive = True Then
    '          lstrFolderPath = lstrExportPath
    '        Else
    '          lstrDocumentPath = String.Format("{0}{1}", lstrExportPath, lobjDocument.ID)
    '        End If

    '        If IO.Directory.Exists(lstrFolderPath) = False Then
    '          Try
    '            IO.Directory.CreateDirectory(lstrFolderPath)
    '          Catch ex As Exception
    '            ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '            ApplicationLogging.WriteLogEntry("Unable to write document to file, filed to create destination directory", TraceEventType.Error, 64723)
    '            Exit Function
    '          End Try
    '        End If

    '      End If

    '      If Args.Archive = True Then
    '        ' Archive the document
    '        lstrFileName = Helper.CleanFile(String.Format("{0}.{1}", lobjDocument.ID, Document.CONTENT_PACKAGE_FILE_EXTENSION), "^")
    '        lstrArchivePath = String.Format("{0}{1}", lstrFolderPath, lstrFileName)
    '        lobjDocument.Archive(lstrArchivePath, True)
    '      End If

    '      If Args.GenerateCDF = True AndAlso Args.Archive = False AndAlso ExportAsArchive = True Then
    '        ' Archive the document
    '        lstrFileName = Helper.CleanFile(String.Format("{0}.{1}", lobjDocument.ID, Document.CONTENT_PACKAGE_FILE_EXTENSION), "^")
    '        lstrArchivePath = String.Format("{0}{1}", lstrFolderPath, lstrFileName)
    '        lobjDocument.Archive(lstrArchivePath, True)
    '      ElseIf Args.GenerateCDF = True Then
    '        ' Serialize the document
    '        lstrFileName = Helper.CleanFile(String.Format("{0}.{1}", lobjDocument.ID, Document.CONTENT_DEFINITION_FILE_EXTENSION), "^")
    '        lobjDocument.Save(String.Format("{0}\{1}", lstrDocumentPath, lstrFileName))
    '      End If

    '      If lstrArchivePath.Length > 0 Then
    '        lobjDocument.SerializationPath = lstrArchivePath
    '      End If

    '    End If

    '    If Args.ErrorMessage.Length = 0 Then
    '      sender.OnDocumentExported(New DocumentExportedEventArgs(lobjDocument, Now, Args.ErrorMessage, Args.Worker))
    '      Return True
    '    Else
    '      sender.OnDocumentExportMessage(New WriteMessageArgs(Args.ErrorMessage))
    '      Return False
    '    End If

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function


    '<Obsolete("This method is no longer supported as of CTS 3.5.", True)>
    'Protected Function ExportDocuments(ByVal sender As IDocumentExporter, ByVal Args As ExportDocumentsEventArgs, ByVal lpExportDocumentCallBack As ExportDocumentCallback) As Boolean
    '  'ByVal sender As Object, ByVal Args As ExportFolderArgs, ByVal lpProvider As CProvider, ByVal lpCallBack As EventForwarderCallback
    '  Try

    '    'Dim lblnReturn As Boolean = True
    '    'Dim lblnFinalReturn As Boolean = True
    '    'Dim lobjFolderContents As FolderContents = Args.FolderContents

    '    'Dim lintSuccessCounter As Integer
    '    'Dim lblnIsEvaluation As Boolean = LicenseRegister.IsEvaluation

    '    'If Args.Worker IsNot Nothing Then
    '    '  '  This will be a background process
    '    '  Dim lobjWorker As System.ComponentModel.BackgroundWorker = Args.Worker
    '    '  If lobjWorker.CancellationPending = True Then
    '    '    Args.DoWorkEventArgs.Cancel = True
    '    '    Return True
    '    '  End If

    '    '  Dim lobjExportDocumentArgs As ExportDocumentEventArgs

    '    '  For Each lobjFolderContent As FolderContent In lobjFolderContents

    '    '    If lobjWorker.CancellationPending = True Then
    '    '      Args.DoWorkEventArgs.Cancel = True
    '    '      Return True
    '    '    End If

    '    '    lobjExportDocumentArgs = New ExportDocumentEventArgs(lobjFolderContent, lobjWorker, Args.Transformation)
    '    '    lobjExportDocumentArgs.GenerateCDF = True

    '    '    lobjExportDocumentArgs.GetRecord = Args.GetRecord
    '    '    lobjExportDocumentArgs.RecordContentSource = Args.RecordContentSource
    '    '    lblnReturn = lpExportDocumentCallBack.Invoke(lobjExportDocumentArgs)
    '    '    If lblnReturn = False Then
    '    '      lblnFinalReturn = False
    '    '      Args.ErrorMessage = lobjExportDocumentArgs.ErrorMessage
    '    '      ApplicationLogging.WriteLogEntry(Args.ErrorMessage, TraceEventType.Error)
    '    '    Else
    '    '      ' Write a success entry in the log
    '    '      ApplicationLogging.WriteLogEntry( _
    '    '        String.Format("Successfully exported document '{0}'.", _
    '    '                      lobjExportDocumentArgs.Id), _
    '    '        TraceEventType.Information, EXPORT_DOCUMENT_SUCCESS)
    '    '      TransactionCounter.Instance.Increment()
    '    '      lintSuccessCounter += 1
    '    '      If lblnIsEvaluation AndAlso lintSuccessCounter = 100 Then
    '    '        Exit For
    '    '      End If
    '    '    End If

    '    '  Next

    '    'Else
    '    '  '  This will be a foreground process
    '    '  Dim lobjExportDocumentArgs As ExportDocumentEventArgs
    '    '  For Each lobjFolderContent As FolderContent In lobjFolderContents

    '    '    lobjExportDocumentArgs = New ExportDocumentEventArgs(lobjFolderContent)
    '    '    lobjExportDocumentArgs.GenerateCDF = True

    '    '    lblnReturn = CType(Me, IDocumentExporter).ExportDocument(lobjFolderContent.ID)
    '    '    If lblnReturn = False Then
    '    '      lblnFinalReturn = False
    '    '      Args.ErrorMessage = lobjExportDocumentArgs.ErrorMessage
    '    '      ApplicationLogging.WriteLogEntry(Args.ErrorMessage, TraceEventType.Error)
    '    '    Else
    '    '      ' Write a success entry in the log
    '    '      ApplicationLogging.WriteLogEntry( _
    '    '        String.Format("Successfully exported document '{0}'.", _
    '    '                      lobjExportDocumentArgs.Id), _
    '    '        TraceEventType.Information, EXPORT_DOCUMENT_SUCCESS)
    '    '      TransactionCounter.Instance.Increment()
    '    '      lintSuccessCounter += 1
    '    '      If lblnIsEvaluation AndAlso lintSuccessCounter = 100 Then
    '    '        Exit For
    '    '      End If
    '    '    End If
    '    '  Next

    '    'End If

    '    'Return lblnFinalReturn

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::ExportDocuments", Me.GetType.Name))
    '  End Try
    'End Function

    Protected Function GetBooleanProviderProperty(ByVal lpPropertyName As String) As Boolean

      Try

        If ProviderProperties.Contains(lpPropertyName) Then

          Select Case ProviderProperties(lpPropertyName).PropertyValue.ToString.ToLower

            Case "false"
              Return False

            Case "true"
              Return True

            Case Else
              Return True
          End Select

        Else
          ' Assume True
          Return True
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Protected Function GetPropertyValue(ByVal lpPropertyName As String) As Object
      Try
        If Me.ProviderProperties(lpPropertyName) IsNot Nothing Then
          Return Me.ProviderProperties(lpPropertyName).PropertyValue
        Else
          Return String.Empty
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::GetPropertyValue('{1}'",
                                      Me.GetType.Name, lpPropertyName))
        Return String.Empty
      End Try

    End Function

    Protected Sub RefreshPropertyFromProviderProperties(lpPropertyName As String)
      Try
        ' Get a reference to the current property
        Dim lobjPropertyInfo As PropertyInfo = Me.GetType.GetProperty(lpPropertyName)

        ' Get a reference to the provider property
        Dim lobjProviderPropertyValue As Object

        ' If the provider property exists and has a value then get it, otherwise
        ' just exit the function since we will have no basis for setting the value.
        If ProviderProperties.Contains(lpPropertyName) AndAlso ProviderProperties(lpPropertyName).HasValue Then
          lobjProviderPropertyValue = ProviderProperties(lpPropertyName).Value
        Else
          Exit Sub
        End If

        ' If we have a property reference we can write to then we can keep going.
        If lobjPropertyInfo IsNot Nothing AndAlso lobjPropertyInfo.CanWrite Then
          lobjPropertyInfo.SetValue(Me, lobjProviderPropertyValue, Nothing)
          '' Get the current value of the property.
          'Dim lobjCurrentPropertyValue As Object = lobjPropertyInfo.GetValue(Me, Nothing)

          '' Check the current value of the property.
          '' If it already has a value then leave it alone, otherwise set it from the provider property.
          'If TypeOf lobjCurrentPropertyValue Is String Then
          '  If String.IsNullOrEmpty(lobjCurrentPropertyValue) Then
          '    lobjPropertyInfo.SetValue(Me, lobjProviderPropertyValue.ToString, Nothing)
          '  End If
          'Else
          '  If lobjCurrentPropertyValue IsNot Nothing Then
          '    lobjPropertyInfo.SetValue(Me, lobjProviderPropertyValue, Nothing)
          '  End If
          'End If

        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub SetPropertyRequired(ByVal lpPropertyName As String, ByVal lpValue As Boolean)
      Try
        Me.ProviderProperties(lpPropertyName)?.SetRequired(lpValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub SetPropertyVisibility(ByVal lpPropertyName As String, ByVal lpVisible As Boolean)
      Try
        Me.ProviderProperties(lpPropertyName)?.SetVisible(lpVisible)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub SetPropertySequence(ByVal lpPropertyName As String, ByVal lpSequenceNumber As Integer)
      Try
        Me.ProviderProperties(lpPropertyName)?.SetSequenceNumber(lpSequenceNumber)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub SetState(lpState As ProviderConnectionState)
      Try
        If lpState <> State Then
          Dim lenuOriginalState As ProviderConnectionState = State
          menuState = lpState
          RaiseEvent ConnectionStateChanged(Me, New ConnectionStateChangedEventArgs(lpState, lenuOriginalState))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Function GetProviderPath() As String
      Try
        Return Me.GetType.Assembly.Location
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::GetProviderPath",
                                      Me.GetType.Name))
        Return ""
      End Try

    End Function

    Private Sub InitializeBaseInformation()
      Try
        'ValidateLicense()
        AddProperties()
        MobjBackgroundWorker.WorkerReportsProgress = True
        MobjBackgroundWorker.WorkerSupportsCancellation = True
        'InitializeImageSet()
      Catch ex As Exception
        mblnIsInitialized = False
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Private Sub InitializeImageSet()
    '  Try
    '    ImageSet = New ImageSet(My.Resources.Initialized, My.Resources.Uninitialized, My.Resources.Unavailable)
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    'Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

#End Region

#Region "Event Handlers"

    Private Sub MobjProviderProperties_CollectionChanged(sender As Object, e As Specialized.NotifyCollectionChangedEventArgs) Handles MobjProviderProperties.CollectionChanged
      Try
        If e.Action = NotifyCollectionChangedAction.Add Then
          For Each lobjNewProperty As ProviderProperty In e.NewItems
            lobjNewProperty.SetParentProvider(Me)
          Next
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub MobjProviderProperties_ItemValueChanged(ByVal sender As Object, ByRef e As Arguments.ProviderPropertyValueChangedEventArgs) Handles MobjProviderProperties.ItemValueChanged
      ' Raise the event to the provider
      RaiseEvent ProviderProperty_ValueChanged(Me, e)
    End Sub

    Private Sub MobjProviderProperties_ProviderPropertyChanged(sender As Object, e As ComponentModel.PropertyChangedEventArgs) Handles MobjProviderProperties.ProviderPropertyChanged
      Try
        RaiseEvent ProviderPropertyChanged(sender, e)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Provider Identification"

    Private Sub AddProperties()
      ' BMK AddProperties
      ' Add the properties here that you want to show up in the 'Create Data Source' dialog.
      Dim lstrDescriptionBuilder As New StringBuilder

      Try
        If SupportsInterface(ProviderClass.DocumentExporter) Then
          ProviderProperties.AddRange(DefineStandardProviderProperties(Providers.ProviderClass.DocumentExporter))
        End If

        If SupportsInterface(ProviderClass.DocumentImporter) Then
          ProviderProperties.AddRange(DefineStandardProviderProperties(Providers.ProviderClass.DocumentImporter))

          If SupportsInterface(ProviderClass.Classification) Then
            ProviderProperties.AddRange(DefineStandardProviderProperties(Providers.ProviderClass.Classification))
          End If

          ' Add SetSystemProperties
          Dim lobjActionProperty As New ActionProperty(ACTION_SET_SYSTEM_PROPERTIES, False, Core.PropertyType.ecmBoolean,
                                                  "Determines whether or not key system properties will be set on the document.") With {
            .Value = False
                                                  }
          ActionProperties.Add(lobjActionProperty)

        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::AddProperties", Me.GetType.Name))
      End Try

    End Sub

    ' <Added by: Ernie at: 2/20/2013-10:37:43 AM on machine: ERNIE-THINK>

    ''' <summary>
    '''     Used to get a collection of all the currently defined standard base provider properties.
    ''' </summary>
    ''' <returns>
    '''     A Ecmg.Providers.ProviderProperties value...
    ''' </returns>
    Public Shared Function GetAllStandardProviderProperties() As ProviderProperties
      Try
        Dim lobjReturnProperties As New ProviderProperties

        lobjReturnProperties.AddRange(DefineStandardProviderProperties(Providers.ProviderClass.DocumentExporter))
        lobjReturnProperties.AddRange(DefineStandardProviderProperties(Providers.ProviderClass.DocumentImporter))
        lobjReturnProperties.AddRange(DefineStandardProviderProperties(Providers.ProviderClass.Classification))

        Return lobjReturnProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    '''     Used the standard set of base provider properties.
    ''' </summary>
    ''' <returns>
    '''     A Ecmg.Providers.ProviderProperties value...
    ''' </returns>
    Private Shared Function DefineStandardProviderProperties(lpInterface As ProviderClass) As ProviderProperties
      Try
        Dim lstrDescriptionBuilder As New StringBuilder
        Dim lobjReturnProperties As New ProviderProperties

        Select Case lpInterface
          Case Providers.ProviderClass.DocumentExporter
            ' Add the 'ExportPath' property
            lobjReturnProperties.Add(New ProviderProperty(EXPORT_PATH, GetType(System.String), True, , ,
              "The default path to export documents to."))

            ' Add the 'ExportAsArchive' property
            lobjReturnProperties.Add(New ProviderProperty(EXPORT_AS_ARCHIVE, GetType(System.Boolean), False, True, ,
              "Specifies whether or not to export documents to a content package file by default."))

            ' Add the 'ExportVersionScope' property
            lobjReturnProperties.Add(New ProviderProperty(EXPORT_VERSION_SCOPE, GetType(VersionScopeEnum), False,
              VersionScopeEnum.AllVersions.ToString, , "Specifies the scope of versions to export.", True))

            ' Add the 'ExportMaxVersions' property
            lobjReturnProperties.Add(New ProviderProperty(EXPORT_MAX_VERSIONS, GetType(Integer), False, 0, ,
              "If applicable to the ExportVersionScope, sets the maximum number of versions to export."))

            ' <Added by: Ernie at: 3/1/2012-8:18:56 AM on machine: ERNIE-M4400>
            ' Added in response to FogBugz case 245 from Duke Energy.
            ' They had cases where they had intentionally stored zero length files in FileNet Content Services.
            ' Add the 'AllowZeroLengthContent' property
            lobjReturnProperties.Add(New ProviderProperty(ALLOW_ZERO_LENGTH_CONTENT, GetType(System.Boolean), False, False, ,
              "Specifies whether or not to allow files with no size (0 bytes) to be exported without raising a ZeroLengthContentException."))
            ' </Added by: Ernie at: 3/1/2012-8:18:56 AM on machine: ERNIE-M4400>

          Case Providers.ProviderClass.DocumentImporter
            ' <Removed by: Ernie at: 7/24/2014-2:35:16 PM on machine: ERNIE-THINK>
            '             '' Add the 'ImportPath' property
            '             'lobjReturnProperties.Add(New ProviderProperty(IMPORT_PATH, GetType(System.String), True))
            ' </Removed by: Ernie at: 7/24/2014-2:35:16 PM on machine: ERNIE-THINK>

          Case Providers.ProviderClass.Classification
            ' Add the 'EnforceClassificationCompliance' property
            ' This should be used to specify whether or not to check for invalid properties in the document.
            lstrDescriptionBuilder = New StringBuilder
            lstrDescriptionBuilder.Append("This should be used to specify whether or not to check for invalid properties in the document.")

            lobjReturnProperties.Add(New ProviderProperty(ENFORCE_CLASSIFICATION_COMPLIANCE, GetType(Boolean), False, True, , lstrDescriptionBuilder.ToString))

            ' Add the 'LogInvalidPropertyRemovals' property
            ' This should be used to specify whether or not to write log entries for each invalid property that is removed
            ' NOTE: This property will be ignored if 'DisableInvalidPropertyChecks' is set to true.
            lstrDescriptionBuilder = New StringBuilder
            lstrDescriptionBuilder.AppendLine("This should be used to specify whether or not to write log entries for each invalid property that is removed.")
            lstrDescriptionBuilder.AppendLine("This property will be ignored if 'DisableInvalidPropertyChecks' is set to true.")
            lobjReturnProperties.Add(New ProviderProperty(LOG_INVALID_PROPERTY_REMOVALS, GetType(Boolean), False, False, , lstrDescriptionBuilder.ToString))

        End Select

        Return lobjReturnProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function
    ' </Added by: Ernie at: 2/20/2013-10:37:43 AM on machine: ERNIE-THINK>

    Private Function GetBooleanPropertyValue(ByVal lpPropertyName As String) As Boolean
      Try
        If ProviderProperties.Contains(lpPropertyName) Then
          Select Case ProviderProperties(lpPropertyName).PropertyValue.ToString.ToLower
            Case "false"
              Return False
            Case "true"
              Return True
            Case Else
              Return True
          End Select
        Else
          ' Assume True
          Return True
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub SetBooleanPropertyValue(ByVal lpPropertyName As String, ByVal lpValue As Boolean)
      Try
        If ProviderProperties.Contains(lpPropertyName) = False Then
          Throw New Exceptions.PropertyDoesNotExistException(
           String.Format("Unable to set value for '{0}', the property does not exist.", lpPropertyName), lpPropertyName)
        End If

        If ProviderProperties(lpPropertyName).PropertyType.Name = "Boolean" Then
          ProviderProperties(lpPropertyName).PropertyValue = lpValue
        Else
          Throw New Exceptions.InvalidPropertyTypeException(
            String.Format("Property '{0}' is type '{1}', a boolean property is expected.",
                          lpPropertyName, ProviderProperties(lpPropertyName).PropertyType.Name))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub SetStringPropertyValue(ByVal lpPropertyName As String, ByVal lpValue As String)
      Try
        If ProviderProperties.Contains(lpPropertyName) = False Then
          Throw New Exceptions.PropertyDoesNotExistException(
           String.Format("Unable to set value for '{0}', the property does not exist.", lpPropertyName), lpPropertyName)
        End If

        If ProviderProperties(lpPropertyName).Type = PropertyType.ecmString Then
          ProviderProperties(lpPropertyName).PropertyValue = lpValue
        Else
          Throw New Exceptions.InvalidPropertyTypeException(
            String.Format("Property '{0}' is type '{1}', a string property is expected.",
                          lpPropertyName, ProviderProperties(lpPropertyName).Type.ToString()))
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IProvider Implementation"

    ''' <summary>
    ''' The connection string used to create the object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shadows Property ConnectionString() As String Implements IProvider.ConnectionString
      Get
        Return MyBase.ConnectionString
      End Get
      Set(ByVal value As String)
        MyBase.ConnectionString = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the folder delimiter used by a specific repository.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride ReadOnly Property FolderDelimiter() As String Implements IProvider.FolderDelimiter

    ''' <summary>
    ''' Gets a value specifying whether or 
    ''' not the repository expects a leading 
    ''' delimiter for all folder paths.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public MustOverride ReadOnly Property LeadingFolderDelimiter() As Boolean Implements IProvider.LeadingFolderDelimiter

    ''' <summary>
    ''' Gets the Provider Class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ProviderClass() As ProviderClass Implements IProvider.ProviderClass
      Get
        Return menuProviderClass
      End Get
    End Property

    ''' <summary>
    ''' Gets the Provider System.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable ReadOnly Property System() As ProviderSystem Implements IProvider.System
      Get
        Return mobjProviderSystem
      End Get
    End Property

    ''' <summary>
    ''' Connects to the repository.
    ''' </summary>
    ''' <remarks></remarks>
    Public Overridable Sub Connect() Implements IProvider.Connect
      ' Throw New NotImplementedException("'Connect' must be implemented in an inherited class")
      Try
        If IsConnected = False Then
          If ContentSource IsNot Nothing Then
            Connect(ContentSource)
            If IsConnected Then
              SetState(ProviderConnectionState.Connected)
            Else
              SetState(ProviderConnectionState.Disconnected)
            End If
          Else
            Throw New NotImplementedException("'Connect()' must be implemented in an inherited class")
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        SetState(ProviderConnectionState.Unavailable)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Connects to the repository.
    ''' </summary>
    ''' <param name="ConnectionString">The connection string used for the connection</param>
    ''' <remarks></remarks>
    Public Overridable Sub Connect(ByVal ConnectionString As String) Implements IProvider.Connect
      Throw New NotImplementedException("'Connect(ConnectionString)' must be implemented in an inherited class")
    End Sub

    ''' <summary>
    ''' Connects to the repository.
    ''' </summary>
    ''' <param name="ContentSource"></param>
    ''' <remarks></remarks>
    Public Overridable Sub Connect(ByVal ContentSource As ContentSource) Implements IProvider.Connect
      Throw New NotImplementedException("'Connect(ContentSource)' must be implemented in an inherited class")
    End Sub

    Public Sub Connect(ByVal Connection As IRepositoryConnection) Implements IProvider.Connect
      Try
        If TypeOf Connection Is ContentSource Then
          Connect(CType(Connection, ContentSource))
        Else
          Throw New InvalidOperationException("Connection is not a ContentSource")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Function ReadyToGetAvailableValues(lpProviderProperty As ProviderProperty) As Boolean _
      Implements IProvider.ReadyToGetAvailableValues
      Try

#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpProviderProperty)
#Else
        If lpProviderProperty Is Nothing Then
          Throw New ArgumentNullException
        End If
#End If

        If lpProviderProperty.SupportsValueList = False Then
          Throw New Exceptions.PropertyDoesNotSupportValueListException(lpProviderProperty)
        End If

        If String.Equals(lpProviderProperty.PropertyType.Name, "Boolean") Then
          Return True
        End If

        If String.Equals(lpProviderProperty.PropertyType.Name, "VersionScopeEnum") Then
          Return True
        End If

        Return False

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function GetAvailableValues(ByVal lpProviderProperty As ProviderProperty) As IEnumerable(Of String) _
      Implements IProvider.GetAvailableValues
      Try

        Dim lobjReturnValues As New List(Of String)

        If lpProviderProperty.SupportsValueList = False Then
          Throw New Exceptions.PropertyDoesNotSupportValueListException(lpProviderProperty)
        End If

        If String.Equals(lpProviderProperty.PropertyType.Name, "Boolean") Then
          lobjReturnValues.Add("False")
          lobjReturnValues.Add("True")
        End If

        If String.Equals(lpProviderProperty.PropertyType.Name, "VersionScopeEnum") Then
          lobjReturnValues.AddRange([Enum].GetNames(GetType(VersionScopeEnum)))
        End If

        Return (lobjReturnValues)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Attempts to return an object of the specified interface
    ''' </summary>
    ''' <param name="lpInterface"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GetInterface(ByVal lpInterface As ProviderClass) As Object Implements IProvider.GetInterface
      Try

        If SupportsInterface(lpInterface) = False Then
          'Throw New InvalidOperationException(String.Format("The provider '{0}' does not support the '{1}' interface.", _
          'Me.Name, lpInterface.ToString))
          Throw New Exceptions.ProviderInterfaceNotImplementedException(Me, lpInterface)
        End If

        Select Case lpInterface
          Case Providers.ProviderClass.BasicContentServices
            Return CType(Me, IBasicContentServicesProvider)

          Case Providers.ProviderClass.Classification
            Return CType(Me, IClassification)

          Case Providers.ProviderClass.Explorer
            Return CType(Me, IExplorer)

          Case Providers.ProviderClass.DocumentExporter
            Return CType(Me, IDocumentExporter)

          Case Providers.ProviderClass.DocumentImporter
            Return CType(Me, IDocumentImporter)

          'Case Providers.ProviderClass.ProviderImages
          '  Return CType(Me, IProviderImages)

          Case Providers.ProviderClass.Copy
            Return CType(Me, ICopy)

          Case Providers.ProviderClass.Delete
            Return CType(Me, IDelete)

          Case Providers.ProviderClass.Rename
            Return CType(Me, IRename)

          Case Providers.ProviderClass.Version
            Return CType(Me, IVersion)

          Case Providers.ProviderClass.File
            Return CType(Me, IFile)

          Case Providers.ProviderClass.CustomObjectClassification
            Return CType(Me, ICustomObjectClassification)

          Case Providers.ProviderClass.FolderClassification
            Return CType(Me, IFolderClassification)

          Case Providers.ProviderClass.FolderManager
            Return CType(Me, IFolderManager)

          Case Providers.ProviderClass.SecurityClassification
            Return CType(Me, ISecurityClassification)

          Case Providers.ProviderClass.Create
            Return CType(Me, ICreate)

          Case Providers.ProviderClass.UpdateProperties
            Return CType(Me, IUpdateProperties)

          Case Providers.ProviderClass.UpdatePermissions
            Return CType(Me, IUpdatePermissions)

          Case Providers.ProviderClass.RecordsManager
            Throw New NotImplementedException("Unable to return IRecordsManager interface object from within Ecmg.Cts")

          Case Providers.ProviderClass.ChoiceListExporter
            ' Throw New NotImplementedException("Unable to return IChoiceListExporter interface object from within Ecmg.Cts")
            Return CType(Me, IChoiceListExporter)

          Case Providers.ProviderClass.ChoiceListImporter
            ' Throw New NotImplementedException("Unable to return IChoiceListImporter interface object from within Ecmg.Cts")
            Return CType(Me, IChoiceListImporter)

          Case Providers.ProviderClass.SQLPassThroughSearch
            If Me.Search IsNot Nothing Then
              Return CType(Me.Search, ISQLPassThroughSearch)
            Else
              Return Nothing
            End If

          Case Providers.ProviderClass.LinkManager
            Return CType(Me, ILinkManager)

          'Case ProviderClass.AnnotationExporter
          '  Return CType(Me, IAnnotationExporter)

          'Case ProviderClass.AnnotationImporter
          '  Return CType(Me, IAnnotationImporter)

          Case ProviderClass.DocumentClassImporter
            Return CType(Me, IDocumentClassImporter)

          Case ProviderClass.RepositoryDiscovery
            Return CType(Me, IRepositoryDiscovery)

          Case ProviderClass.FolderExporter
            Return CType(Me, IFolderExporter)

          Case ProviderClass.FolderImporter
            Return CType(Me, IFolderImporter)

          Case ProviderClass.CustomObjectExporter
            Return CType(Me, ICustomObjectExporter)

          Case ProviderClass.CustomObjectImporter
            Return CType(Me, ICustomObjectImporter)

          Case Else
            Return Nothing

        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IAuthorization Implementation"

    Public Overridable Function Login(ByVal lpUserName As String, ByVal lpPassword As String) As Boolean Implements IAuthorization.Login
      Throw New NotImplementedException("'Login' must be implemented in an inherited class")
    End Function

    Public Overridable Function Logout() As Boolean Implements IAuthorization.Logout
      Throw New NotImplementedException("'Logout' must be implemented in an inherited class")
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
        Return PROVIDER_INFORMATION_FILE_EXTENSION
      End Get
    End Property

    Public Function Deserialize(ByVal lpFilePath As String, Optional ByRef lpErrorMessage As String = "") As Object Implements ISerialize.Deserialize
      Try
        Return Serializer.Deserialize.XmlFile(lpFilePath, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::Deserialize('{1}')", Me.GetType.Name, lpFilePath))
        lpErrorMessage = Helper.FormatCallStack(ex)
        Return Nothing
      End Try
    End Function

    Public Function DeSerialize(ByVal lpXML As System.Xml.XmlDocument) As Object Implements ISerialize.DeSerialize
      Try
        Return Serializer.Deserialize.XmlString(lpXML.OuterXml, Me.GetType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::Deserialize(lpXML)", Me.GetType.Name))
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
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::Serialize('{1}')", Me.GetType.Name, lpFilePath))
      End Try

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
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::Serialize('{1}', '{2}', '{3}')",
                                      Me.GetType.Name, lpFilePath, lpWriteProcessingInstruction, lpStyleSheetPath))
      End Try

    End Sub

    Public Function Serialize() As System.Xml.XmlDocument Implements ISerialize.Serialize
      Try
        Return Serializer.Serialize.Xml(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::Serialize", Me.GetType().Name))
        Return Nothing
      End Try
    End Function

    Public ReadOnly Property SerializationPath() As String
      Get
        Return mstrSerializationPath
      End Get
    End Property

    Public Function SetSerializationPath(ByVal lpSerializationPath As String) As String

      mstrSerializationPath = lpSerializationPath
      Return SerializationPath

    End Function

    Public Function ToXmlString() As String Implements ISerialize.ToXmlString
      Try
        Return Serializer.Serialize.XmlString(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::ToString", Me.GetType.Name))
        Return ""
      End Try
    End Function

#End Region

#Region "Background Worker Event Handlers"

    Private Sub MobjBackgroundWorker_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MobjBackgroundWorker.Disposed
      RaiseEvent BackgroundWorker_Disposed(sender, e)
    End Sub

    Private Sub MobjBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles MobjBackgroundWorker.DoWork
      RaiseEvent BackgroundWorker_DoWork(sender, e)
    End Sub

    Private Sub MobjBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles MobjBackgroundWorker.ProgressChanged
      Try
        If e.ProgressPercentage > -1 Then
          RaiseEvent BackgroundWorker_ProgressChanged(sender, e)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub MobjBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles MobjBackgroundWorker.RunWorkerCompleted
      RaiseEvent BackgroundWorker_RunWorkerCompleted(sender, e)
    End Sub

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      Return Nothing
    End Function

    Public Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml

      'Dim value As Object
      'Dim lobjValue As Object
      Dim lobjObject As New ProviderSystem ' ISerialize

      Dim lstrName As String
      Dim lstrType As String
      Dim lstrCompany As String
      Dim lstrProductName As String
      Dim lstrProductVersion As String
      Dim lstrPath As String

      'lobjObject = New ProviderSystem

      Try

        With reader
          .Read() ' move past container
          .ReadStartElement("ProviderSystem")
          lstrName = reader.ReadElementString("Name")
          lstrType = reader.ReadElementString("Type")
          lstrCompany = reader.ReadElementString("Company")
          lstrProductName = reader.ReadElementString("ProductName")
          lstrProductVersion = reader.ReadElementString("ProductVersion")
          lstrPath = reader.ReadElementString("Path")
          .ReadEndElement()
        End With

        With Me
          .ProviderSystem.Company = lstrCompany
          .ProviderSystem.Name = lstrName
          .ProviderSystem.ProductName = lstrProductName
          .ProviderSystem.ProductVersion = lstrProductVersion
          .SetSerializationPath(lstrPath)
        End With
        '   Beep()

        'While reader.NodeType <> XmlNodeType.EndElement
        '  '    reader.ReadStartElement("Provider")
        '  '    'reader.ReadStartElement()
        '  '    Dim key As String = reader.ReadElementString("Key")
        '  '    'value = reader.ReadElementString("value", NS)
        '  '    lobjClassificationProperty = New ClassificationProperty
        '  lobjXMLDocument.LoadXml(reader.ReadInnerXml)
        '  lobjObject.Deserialize(lobjXMLDocument)
        '  '    Select Case lobjXMLDocument.FirstChild.Name
        '  '      Case "ClassificationProperty"
        '  '        lobjObject = New ClassificationProperty
        '  '      Case "ECMProperty"
        '  '        lobjObject = New ECMProperty
        '  '      Case Else
        '  '        lobjObject = New ECMProperty
        '  '    End Select
        '  '    lobjObject.Deserialize(lobjXMLDocument)
        '  '    'lobjValue = reader.ReadElementContentAsObject()
        '  '    'lobjValue = Ecmg.ECMUtilities.Helper.Deserialize
        '  '    'value = lobjClassificationProperty.Deserialize(lobjXMLDocument)
        '  '    'reader.ReadEndElement()
        '  '    reader.MoveToContent()
        '  '    lobjDictionary.Add(key, lobjObject)
        'End While

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::ReadXml", Me.GetType.Name))
      End Try

    End Sub

    Public Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml

      Try

        With writer

          ' Open the 'ProviderSystem' Element
          .WriteStartElement("ProviderSystem", "")

          ' Write out the 'ProviderSystem'
          '     Write out the 'Name'
          .WriteElementString("Name", ProviderSystem.Name)
          '     Write out the 'Type'
          .WriteElementString("Type", ProviderSystem.Type)
          '     Write out the 'Company'
          .WriteElementString("Company", ProviderSystem.Company)
          '     Write out the 'ProductName'
          .WriteElementString("ProductName", ProviderSystem.ProductName)
          '     Write out the 'ProductVersion'
          .WriteElementString("ProductVersion", ProviderSystem.ProductVersion)
          '     Write out the 'Path'
          .WriteElementString("Path", Me.GetType.Assembly.Location)

          ' Close the 'Provider' Element
          .WriteEndElement()

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::CProvider::WriteXml", Me.GetType.Name))
      End Try

    End Sub

#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
      Try
        If Not Me.disposedValue Then
          If disposing Then
            ' DISPOSETODO: dispose managed state (managed objects).
            If IsConnected Then
              Disconnect()
            End If
            mobjActionProperties = New ActionProperties
            mobjContentSource = Nothing
            mblnHasValidLicense = False
            mobjProviderInformation = Nothing
            menuProviderClass = Nothing
            mobjProviderSystem = Nothing
            mstrSerializationPath = Nothing
            mstrExportPath = Nothing
            'mstrImportPath = Nothing
            MobjProviderProperties = Nothing
            MobjBackgroundWorker = Nothing
            mblnIsConnected = False
            mblnIsInitialized = False
            'mobjImageSet = Nothing
            mintMaxContentCount = -1
            mstrLicenseFailureReason = Nothing
            mblnEnforceClassificationCompliance = False
            mblnLogInvalidPropertyRemovals = Nothing
            mblnClassificationPropertiesInitialized = False
            menuState = ProviderConnectionState.Disconnected
            mblnAllowZeroLengthContent = Nothing
            mobjSelectedFolder = Nothing
            mobjTag = Nothing

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
#End Region

  End Class

End Namespace