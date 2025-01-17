'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports System.Xml.Serialization
Imports Documents.Files
Imports Documents.Network
Imports Documents.Utilities
Imports Newtonsoft.Json


#End Region

Namespace Core

  ''' <summary>Defines the content of a document version.</summary>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Content
    Implements IMetaHolder
    Implements IDisposable

#Region "Public Enumerations"

    ''' <summary>
    ''' Determines the manner in which the content file will be handled
    ''' </summary>
    Public Enum StorageTypeEnum
      ''' <summary>
      ''' The content is not stored internally, but rather is referenced externally through the ContentPath property.
      ''' </summary>
      Reference = 0
      ''' <summary>
      ''' The content is stored internally as a Base64 encoded string in the Data property
      ''' </summary>
      EncodedUnCompressed = 1
      ''' <summary>
      ''' The content is compressed and stored internally as a Base64 encoded string in the Data property
      ''' </summary>
      EncodedCompressed = 2
    End Enum

#End Region

#Region "Class Constants"

    Private Const MAX_FILE_LENGTH As Integer = 239

#End Region

#Region "Class Variables"

    Protected mstrContentPath As String
    Private mobjContentData As New ContentData(Me)
    Private mstrRelativePath As String = String.Empty
    Protected mstrMimeType As String = String.Empty
    Private menuStorageType As StorageTypeEnum = StorageTypeEnum.Reference
    Private menuPreviousStorageType As StorageTypeEnum = StorageTypeEnum.Reference
    Private mstrHash As String = String.Empty
    Private mobjFileSize As FileSize
    Private mstrFileName As String
    Private mobjMetadata As New ECMProperties
    Private mobjVersion As Version
    Private mobjProperties As New ECMProperties
    Private mblnGetAvailableMetadata As Nullable(Of Boolean) = Nothing
    Private menuProtocol As ProtocolEnum = ProtocolEnum.file

    ' Used as a reference to the parent document
    Private mobjDocument As Document

#End Region

#Region "Public Properties"

#Region "Read Only Properties"

    <JsonIgnore()>
    Public ReadOnly Property CanRead As Boolean
      Get
        If mobjContentData IsNot Nothing Then
          Return mobjContentData.CanRead
        Else
          Return False
        End If
      End Get
    End Property

    ''' <summary>
    ''' Gets the current path to the content file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property CurrentPath() As String
      Get
        Try
          Return GetCurrentPath()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    <JsonIgnore()>
    Public ReadOnly Property Protocol As ProtocolEnum
      Get
        Try
          Return menuProtocol
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Friend ReadOnly Property ShouldGetAvailableMetadata As Nullable(Of Boolean)
      Get
        Return mblnGetAvailableMetadata
      End Get
    End Property

    ''' <summary>
    ''' Gets the Document to which this Content belongs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property Document() As Document
      Get
        Try
          If mobjVersion IsNot Nothing AndAlso mobjVersion.Document IsNot Nothing Then
            Return mobjVersion.Document
          Else
            'ApplicationLogging.WriteLogEntry("Unable to get Document property of Content object.  The Version and or the Version.Document object was not set.", TraceEventType.Warning)
            Return mobjDocument
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the Version to which this Content belongs
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property Version() As Version
      Get
        Return mobjVersion
      End Get
    End Property

    ''' <summary>
    '''     Determines whether the content file includes a file name extension.
    ''' </summary>
    <JsonIgnore()>
    Public ReadOnly Property HasExtension As Boolean
      Get
        Try
          If Not String.IsNullOrEmpty(mstrContentPath) Then
            Return Path.HasExtension(mstrContentPath)
          Else
            Return Path.HasExtension(FileName)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the file extension of the ContentPath
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileExtension() As String
      Get
        Try
          If Not String.IsNullOrEmpty(mstrContentPath) Then
            Return GetFileExtension(mstrContentPath)
          Else
            Return GetFileExtension(FileName)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
        'Return GetLastPart(mstrContentPath, ".")
      End Get
    End Property

#End Region

    ''' <summary>
    ''' Gets or sets the path relative to the cdf file for the content element.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public Property RelativePath() As String
      Get
        Return mstrRelativePath
      End Get
      Set(ByVal value As String)
        mstrRelativePath = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the fully qualified file name for the content element.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Property ContentPath() As String
      Get
        Try

          'If we are encoded there is no need to check if file exists, just return content path as it is in the CDF 
          If (menuStorageType = StorageTypeEnum.EncodedCompressed Or menuStorageType = StorageTypeEnum.EncodedUnCompressed) Then
            Return mstrContentPath
          End If

          If Protocol = Network.ProtocolEnum.file Then
            If IO.File.Exists(mstrContentPath) Then
              Return mstrContentPath
            Else
              Return GetCurrentPath()
            End If
          Else
            Return mstrContentPath
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal Value As String)
        Try
          mstrContentPath = Value
          ' Save the FileName property as well
          mstrFileName = Path.GetFileName(mstrContentPath)
          mobjContentData.ContentPath = Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets the file name of the ContentPath
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FileName() As String
      Get
        Return mstrFileName
      End Get
      Set(ByVal value As String)
        If Helper.IsDeserializationBasedCall Then
          mstrFileName = value
        Else
          Throw New InvalidOperationException("Although 'FileName' is a public property, set operations are not allowed.  Treat as read-only.")
        End If
      End Set
    End Property

    ''' <summary>
    ''' Determines the manner in which the content file will be handled.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public Overridable Property StorageType() As StorageTypeEnum
      Get
        Return menuStorageType
      End Get
      Set(ByVal value As StorageTypeEnum)
        PreviousStorageType = menuStorageType
        menuStorageType = value
        mobjContentData.StorageType = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the content data as a Base64 Encoded string
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public Property Data() As ContentData
      Get
        Return mobjContentData
      End Get
      Set(ByVal value As ContentData)
        Try
          mobjContentData = value
          mobjContentData.SetContent(Me)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets a FileSize object for the contained content element.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FileSize() As FileSize
      Get
        If mobjFileSize Is Nothing Then
          mobjFileSize = New FileSize(FileContentSize)
        End If
        Return mobjFileSize
      End Get
      Set(ByVal Value As FileSize)
        If Helper.CallStackContainsMethodName("Deserialize") Then
          mobjFileSize = Value
        Else
          Throw New InvalidOperationException("Set operation not supported for FileSizeProperty")
        End If
      End Set
    End Property

    <XmlElement("Metadata")>
    <JsonProperty(NameOf(Metadata))>
    Public Property ContentProperties As ECMProperties
      Get
        Return mobjMetadata
      End Get
      Set(ByVal value As ECMProperties)
        mobjMetadata = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the available properties for the content element.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The metadata contains the properties persisted inside the actual content file.</remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property Metadata() As IProperties Implements IMetaHolder.Metadata
      Get
        Return mobjMetadata
      End Get
      Set(ByVal value As IProperties)
        mobjMetadata = value
      End Set
    End Property


#End Region

#Region "Private Properties"

    Private Property ArchiveContentPath As String
      Get
        Return mstrContentPath
      End Get
      Set(ByVal value As String)
        Try
          mstrContentPath = value
          ' Save the FileName property as well
          mstrFileName = Path.GetFileName(mstrContentPath)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Private Property PreviousStorageType() As StorageTypeEnum
      Get
        Return menuPreviousStorageType
      End Get
      Set(ByVal value As StorageTypeEnum)
        menuPreviousStorageType = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
    End Sub

    ''' <summary>
    ''' Takes the specified content path and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <remarks>Defaults the StorageType to Reference and attempts to get available metadata</remarks>
    Public Sub New(ByVal lpContentPath As String)
      Me.New(lpContentPath, StorageTypeEnum.Reference)
    End Sub

    Public Sub New(ByVal lpVersion As Version)
      mobjVersion = lpVersion
    End Sub

    Public Sub New(ByVal lpContentPath As String, ByVal lpVersion As Version)
      Me.New(lpContentPath, StorageTypeEnum.Reference, True, lpVersion, String.Empty, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpVersion As Version,
                   ByVal lpAllowZeroLengthContent As Boolean)
      Me.New(lpContentPath, StorageTypeEnum.Reference, True, lpVersion, String.Empty, lpAllowZeroLengthContent)
    End Sub

    ''' <summary>
    ''' Takes the specified content path and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <param name="lpFileName">Sets the file name associated with the content</param>
    ''' <remarks>Defaults the StorageType to Reference and attempts to get available metadata</remarks>
    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpFileName As String)
      ' Me.New(lpContentPath, StorageTypeEnum.Reference)
      Me.New(lpContentPath, StorageTypeEnum.Reference, True, Nothing, lpFileName, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpFileName As String, ByRef lpDocument As Document)
      ' Me.New(lpContentPath, StorageTypeEnum.Reference)
      Me.New(lpContentPath, StorageTypeEnum.Reference, True, Nothing, lpFileName, False, lpDocument)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpVersion As Version,
                   ByVal lpFileName As String)
      Me.New(lpContentPath, StorageTypeEnum.Reference, True, lpVersion, lpFileName, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpVersion As Version,
                   ByVal lpFileName As String,
                   ByVal lpAllowZeroLengthContent As Boolean)
      Me.New(lpContentPath, StorageTypeEnum.Reference, True, lpVersion, lpFileName, lpAllowZeroLengthContent)
    End Sub

    ''' <summary>
    ''' Takes the specified content path and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <param name="lpGetAvailableMetadata">Specifies whether or not to attempt to open the file and read in the available metadata</param>
    ''' <remarks>Defaults the StorageType to Reference</remarks>
    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpGetAvailableMetadata As Boolean,
                   ByVal lpFileName As String)
      Me.New(lpContentPath, StorageTypeEnum.Reference, lpGetAvailableMetadata, Nothing, lpFileName, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpGetAvailableMetadata As Boolean,
                   ByVal lpVersion As Version,
                   ByVal lpFileName As String)
      Me.New(lpContentPath, StorageTypeEnum.Reference, lpGetAvailableMetadata, lpVersion, lpFileName, False)
    End Sub

    ''' <summary>
    ''' Takes the specified content path and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <param name="lpGetAvailableMetadata">Specifies whether or not to attempt to open the file and read in the available metadata</param>
    ''' <remarks>Defaults the StorageType to Reference</remarks>
    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpGetAvailableMetadata As Boolean)
      Me.New(lpContentPath, StorageTypeEnum.Reference, lpGetAvailableMetadata, Nothing, String.Empty, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpGetAvailableMetadata As Boolean,
                   ByVal lpVersion As Version)
      Me.New(lpContentPath, StorageTypeEnum.Reference, lpGetAvailableMetadata, lpVersion, String.Empty, False)
    End Sub

    ''' <summary>
    ''' Takes the specified content path and Storage Type and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <param name="lpStorageType">Specifies whether the content file is referenced or encoded</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpStorageType As StorageTypeEnum)
      Me.New(lpContentPath, lpStorageType, True, Nothing, String.Empty, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpStorageType As StorageTypeEnum,
                   ByVal lpVersion As Version)
      Me.New(lpContentPath, lpStorageType, True, lpVersion, String.Empty, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
               ByVal lpStorageType As StorageTypeEnum,
               ByVal lpVersion As Version,
               ByVal lpAllowZeroLengthContent As Boolean)
      Me.New(lpContentPath, lpStorageType, True, lpVersion, String.Empty, lpAllowZeroLengthContent)
    End Sub

    ''' <summary>
    ''' Takes the specified content path and Storage Type and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <param name="lpStorageType">Specifies whether the content file is referenced or encoded</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpStorageType As StorageTypeEnum,
                   ByVal lpFileName As String)
      Me.New(lpContentPath, lpStorageType, True, Nothing, lpFileName, False)
    End Sub

    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpStorageType As StorageTypeEnum,
                   ByVal lpVersion As Version,
                   ByVal lpFileName As String)
      Me.New(lpContentPath, lpStorageType, True, lpVersion, lpFileName, False)
    End Sub

    ''' <summary>
    ''' Takes the specified content path and Storage Type and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <param name="lpStorageType">Specifies whether the content file is referenced or encoded</param>
    ''' <param name="lpGetAvailableMetadata">Specifies whether or not to attempt 
    ''' to open the file and read in the available metadata</param>
    ''' <param name="lpFileName">The complete file name without folder path, 
    ''' This may be passed as an empty string in which case the value wil be inferred from the content path</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpContentPath As String,
               ByVal lpStorageType As StorageTypeEnum,
               ByVal lpGetAvailableMetadata As Boolean,
               ByVal lpVersion As Version,
               ByVal lpFileName As String,
               ByVal lpAllowZeroLengthContent As Boolean)

      Me.New(lpContentPath, lpStorageType, lpGetAvailableMetadata,
             lpVersion, lpFileName, lpAllowZeroLengthContent, False, Nothing)

    End Sub

    Public Sub New(ByVal lpContentPath As String,
               ByVal lpStorageType As StorageTypeEnum,
               ByVal lpGetAvailableMetadata As Boolean,
               ByVal lpVersion As Version,
               ByVal lpFileName As String,
               ByVal lpAllowZeroLengthContent As Boolean,
               ByRef lpDocument As Document)

      Me.New(lpContentPath, lpStorageType, lpGetAvailableMetadata,
             lpVersion, lpFileName, lpAllowZeroLengthContent, False, lpDocument)

    End Sub
    ''' <summary>
    ''' Takes the specified content path and Storage Type and creates a new Content object
    ''' </summary>
    ''' <param name="lpContentPath">The fully qualified file name for the content file</param>
    ''' <param name="lpStorageType">Specifies whether the content file is referenced or encoded</param>
    ''' <param name="lpGetAvailableMetadata">Specifies whether or not to attempt 
    ''' to open the file and read in the available metadata</param>
    ''' <param name="lpFileName">The complete file name without folder path, 
    ''' This may be passed as an empty string in which case the value wil be inferred from the content path</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpContentPath As String,
                   ByVal lpStorageType As StorageTypeEnum,
                   ByVal lpGetAvailableMetadata As Boolean,
                   ByVal lpVersion As Version,
                   ByVal lpFileName As String,
                   ByVal lpAllowZeroLengthContent As Boolean,
                   ByVal lpAnalyze As Boolean,
                   ByRef lpDocument As Document)
      Try

        'Dim lstrTempContentDirectory As String
        'Dim lstrTempContentPath As String

        mobjDocument = lpDocument

        menuProtocol = Helper.GetProtocol(lpContentPath)

        mobjVersion = lpVersion

        If Protocol = ProtocolEnum.file Then
          '  Check to make sure the specified content path is valid
          If IO.File.Exists(lpContentPath) = False Then
            Throw New Exceptions.InvalidPathException(String.Format("Unable to create Content object, the path '{0}' is invalid.", lpContentPath), lpContentPath)
          End If
        End If

        '' Check for zero byte file before creating the content object so we can more clearly report it.
        '' Ernie Bahr 8/3/2016
        'If lpAllowZeroLengthContent = False AndAlso FileHelper.IsZeroByteFile(lpContentPath) Then
        '  Throw New ZeroLengthContentException(String.Format("Content for file '{0}' has zero length.", lpContentPath))
        'End If

        ContentPath = lpContentPath

        Select Case Protocol
          Case ProtocolEnum.http, ProtocolEnum.https
            mobjContentData.SetStream(Helper.WriteFileToMemoryStream(lpContentPath))
            mstrFileName = IO.Path.GetFileName(lpContentPath)
        End Select

        If lpFileName <> String.Empty Then
          mstrFileName = lpFileName
        End If

        ' Initialize the file size here
        Dim llngFileContentSize As Long = FileContentSize()
        mobjFileSize = New FileSize(llngFileContentSize)

        If llngFileContentSize = 0 AndAlso lpAllowZeroLengthContent = False Then
          Throw New Exceptions.ZeroLengthContentException(String.Format("Content for file '{0}' has zero length.", lpContentPath))
        End If

        StorageType = lpStorageType

        mblnGetAvailableMetadata = lpGetAvailableMetadata

#If Not CTSSILVERLIGHT = 1 Then

        If lpAnalyze Then
          Analyze(Precision.Quick)
        End If

#End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates a new content object with the specified byte array named using the specified content file name
    ''' </summary>
    ''' <param name="lpByteArray">An array of bytes containing the file content</param>
    ''' <param name="lpContentFileName">The simple qualified file name for the content</param>
    ''' <remarks>Attempts to get available metadata</remarks>
    Public Sub New(ByVal lpByteArray As Byte(),
                   ByVal lpContentFileName As String,
                   ByVal lpStorageType As StorageTypeEnum)
      Me.New(lpByteArray, lpContentFileName, lpStorageType, False, False)
    End Sub

    ''' <summary>
    ''' Creates a new content object with the specified byte array named using the specified content file name
    ''' </summary>
    ''' <param name="lpByteArray">An array of bytes containing the file content</param>
    ''' <param name="lpContentFileName">The simple qualified file name for the content</param>
    ''' <param name="lpGetAvailableMetadata">Specifies whether or not to attempt to open the file and read in the available metadata</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpByteArray As Byte(),
                   ByVal lpContentFileName As String,
                   ByVal lpStorageType As StorageTypeEnum,
                   ByVal lpGetAvailableMetadata As Boolean,
                   ByVal lpAllowZeroLengthContent As Boolean)
      Try
        'Dim lobjByteArray As Byte()
        'Dim lintStreamLength As Integer
        'Dim lstrFileName As String = String.Format("{0}{1}", FileHelper.GetSpecialPath(Helper.SpecialPathType.CtsTemp), lpContentFileName)

        'lobjByteArray = New Byte(lpStream.Length) {}
        'lintStreamLength = lpStream.Read(lobjByteArray, 0, lpStream.Length)

        ' Added by Ernie Bahr on 9/27/2011
        ' Validate the byte array
        If lpByteArray.Length = 0 AndAlso lpAllowZeroLengthContent = False Then
          Throw New Exceptions.ZeroLengthContentException(String.Format("Content for file '{0}' has zero length.", lpContentFileName))
        End If
        ' End Added by Ernie Bahr on 9/27/2011

        'EncodeFile(lobjByteArray, lstrFileName)
        'ContentPath = lpContentFileName
        mstrFileName = lpContentFileName

        ' Initialize the file size here
        'mobjFileSize = New FileSize(FileContentSize())
        StorageType = lpStorageType
        mobjContentData.SetStream(lpByteArray)

        'If lpGetAvailableMetadata = True Then
        '  GetAvailableMetadata()
        'End If

        mblnGetAvailableMetadata = lpGetAvailableMetadata

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Create a new content object with the specified named stream object
    ''' </summary>
    ''' <param name="lpStream">A NamedStream object containing the file content</param>
    Public Sub New(lpStream As NamedStream)
      Me.New(lpStream.Stream, lpStream.FileName, StorageTypeEnum.Reference, False, False)
    End Sub

    ''' <summary>
    ''' Create a new content object with the specified stream object using the specified content file name
    ''' </summary>
    ''' <param name="lpStream">A IO.Stream object containing the file content</param>
    ''' <param name="lpContentFileName">The simple qualified file name for the content</param>
    ''' <remarks>Attempts to get available metadata</remarks>
    Public Sub New(ByVal lpStream As IO.Stream,
                   ByVal lpContentFileName As String,
                   ByVal lpStorageType As StorageTypeEnum)
      Me.New(lpStream, lpContentFileName, lpStorageType, False, False)
    End Sub

    ''' <summary>
    ''' Create a new content object with the specified stream object using the specified content file name
    ''' </summary>
    ''' <param name="lpStream">A IO.Stream object containing the file content</param>
    ''' <param name="lpContentFileName">The simple qualified file name for the content</param>
    ''' <param name="lpGetAvailableMetadata">Specifies whether or not to attempt to open the file and read in the available metadata</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpStream As IO.Stream,
                   ByVal lpContentFileName As String,
                   ByVal lpStorageType As StorageTypeEnum,
                   ByVal lpGetAvailableMetadata As Boolean,
                   ByVal lpAllowZeroLengthContent As Boolean)
      Try
        'Dim lobjByteArray As Byte()
        'Dim lintStreamLength As Integer
        'Dim lstrFileName As String = String.Format("{0}{1}", FileHelper.GetSpecialPath(Helper.SpecialPathType.CtsTemp), lpContentFileName)

        'lobjByteArray = New Byte(lpStream.Length) {}
        'lintStreamLength = lpStream.Read(lobjByteArray, 0, lpStream.Length)

        ' Added by Ernie Bahr on 9/27/2011
        ' Validate the incoming stream


#If NET8_0_OR_GREATER Then
        ArgumentNullException.ThrowIfNull(lpStream)
#Else
        If lpStream Is Nothing Then
          Throw New ArgumentNullException
        End If
#End If

        If lpStream.CanSeek Then
          If lpStream.Length = 0 AndAlso lpAllowZeroLengthContent = False Then
            Throw New Exceptions.ZeroLengthContentException(String.Format("Content for file '{0}' has zero length.", lpContentFileName))
          End If
        End If
        ' End Added by Ernie Bahr on 9/27/2011

        'EncodeFile(lobjByteArray, lstrFileName)
        'ContentPath = lpContentFileName
        mstrFileName = lpContentFileName

        ' Initialize the file size here
        'mobjFileSize = New FileSize(FileContentSize())
        StorageType = lpStorageType

        'Set the bytes in the content data object.
        ' <Modified by: Ernie at 3/2/2012-10:57:28 AM on machine: ERNIE-M4400>
        If lpStream.CanSeek Then
          If lpStream.Length > 0 Then
            mobjContentData.SetStream(lpStream)
          End If
        Else
          mobjContentData.SetStream(lpStream)
        End If

        ' </Modified by: Ernie at 3/2/2012-10:57:28 AM on machine: ERNIE-M4400>
        'mobjContentData.SetStream(lpStream)

        'If lpGetAvailableMetadata = True Then
        '  GetAvailableMetadata()
        'End If
        mblnGetAvailableMetadata = lpGetAvailableMetadata

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Gets a collection of properties based on the actual root 
    ''' properties of this object as well as the metadata collection
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CompleteProperties() As ECMProperties
      Return GetContentProperties(True)
    End Function

    ''' <summary>
    ''' Gets a collection of properties based on the actual root properties of this object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Properties() As ECMProperties
      Return GetContentProperties(False)
    End Function

    Friend Sub SetData(ByVal lpStorageType As StorageTypeEnum)
      Try
        Select Case lpStorageType
          '  If the content file is to be encoded
          Case StorageTypeEnum.EncodedCompressed, StorageTypeEnum.EncodedUnCompressed
            '  Only continue if it is either not yet encoded or the storage type has changed
            If Me.Data Is Nothing OrElse Me.Data.Length = 0 OrElse Me.StorageType <> lpStorageType Then
              '  Ensure that the content path is valid
              If IO.File.Exists(ContentPath) Then
                '  Encode the data
                EncodeData(lpStorageType)
              Else
                Throw New Exceptions.InvalidPathException(String.Format("Unable to set data: ContentPath '{0}'", ContentPath), ContentPath)
              End If
            End If
          Case StorageTypeEnum.Reference
            '  Only continue of the storage path has changed
            If Me.Data IsNot Nothing OrElse File.Exists(ContentPath) = False Then
              '  Encode the data
              EncodeData(lpStorageType)
            End If
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Sub EncodeData(ByVal lpStorageType As StorageTypeEnum)
      Try
        Select Case lpStorageType
          Case StorageTypeEnum.EncodedCompressed
            'RKS commented out next line
            ''Me.Data = EncodeFileToString(ContentPath, True)
            '  Delete all the content files
            'DeleteContentFiles()
          Case StorageTypeEnum.EncodedUnCompressed
            'RKS commented out next line
            ''Me.Data = EncodeFileToString(ContentPath, False)
            '  Delete all the content files
            'DeleteContentFiles()
          Case StorageTypeEnum.Reference
            '  Recreate the content file from the previously encoded data
            'Why do this, <data> will always be empty if storagetype is reference???
            'RecreateContentFile()
            ''Me.Data = ""
          Case Else
            Throw New ArgumentException(String.Format("The value '{0}' is not a valid value for a StorageType", lpStorageType), NameOf(lpStorageType))
        End Select

        '  Set the StorageType property to the correct type
        Me.StorageType = lpStorageType

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

#If Not CTSSILVERLIGHT = 1 Then

    Private Sub Analyze(lpPrecision As Precision)
      Try
        If Not CanRead Then
          Throw New InvalidOperationException("Content can't be read.")
        End If

        Dim lobjAnalyzer As Analyzer = Analyzer.Instance
        lobjAnalyzer.Precision = lpPrecision

        Dim lobjResults As IFileTestResults = lobjAnalyzer.Analyze(Me.Data.ToStream())

        Dim lobjPrimaryResult As IFileTestResult = lobjResults.PrimaryResult

        If lobjPrimaryResult IsNot Nothing Then
          'If Not String.IsNullOrEmpty(lobjPrimaryResult.Extension) Then
          '  If Not String.Equals(lobjPrimaryResult.Extension, Me.FileExtension) Then
          '    Me.FileExtension = lobjPrimaryResult.Extension
          '  End If
          'End If
          If Not String.IsNullOrEmpty(lobjPrimaryResult.MimeType) Then
            Me.MIMEType = lobjPrimaryResult.MimeType
          End If

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End If

    Private Function GetFileExtension(ByVal lpFileName As String) As String
      Try
        Dim lstrRawExtension As String = Path.GetExtension(lpFileName)

        If lstrRawExtension.StartsWith("."c) Then
          Return lstrRawExtension.Remove(0, 1)
        Else
          Return lstrRawExtension
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, lpFileName)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetContentProperties(ByVal lpIncludeMetadata As Boolean) As ECMProperties
      Try
        ' Go through the root level content properties and return them in a properties collection
        Dim lobjContentProperties As New ECMProperties From {
          PropertyFactory.Create(PropertyType.ecmString, "RelativePath", Me.RelativePath),
          PropertyFactory.Create(PropertyType.ecmString, "ContentPath", Me.ContentPath),
          PropertyFactory.Create(PropertyType.ecmString, "StorageType", Me.StorageType.ToString),
          PropertyFactory.Create(PropertyType.ecmString, "FileName", Me.FileName),
          PropertyFactory.Create(PropertyType.ecmString, "FileExtension", Me.FileExtension),
          GetExtendedContentProperties()
        }

        If lpIncludeMetadata = True Then
          lobjContentProperties.Add(Metadata)
        End If

        Return lobjContentProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function FileContentSize() As Long
      Try

        If Protocol = ProtocolEnum.file Then
          If (File.Exists(mstrContentPath) = False) Then 'AndAlso ((Data Is Nothing) OrElse (Data.Length = 0)) Then
            If (Data Is Nothing) OrElse (Data.Length = 0) Then
              Return 0
            Else
              Return Data.Length
            End If
          End If

          Dim lobjFileInfo As New IO.FileInfo(mstrContentPath)

          Return lobjFileInfo.Length
        Else
          If (Data Is Nothing) OrElse (Data.Length = 0) Then
            Return 0
          Else
            Return Data.Length
          End If
        End If

      Catch Ex As Exception
        ApplicationLogging.LogException(Ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    Private Function GetCurrentPath() As String
      Try

        If Document IsNot Nothing AndAlso
          Document.CurrentPath IsNot Nothing AndAlso
            Document.CurrentPath <> String.Empty Then

          Dim lstrCurrentPath As String = String.Format("{0}\{1}\{2}", IO.Path.GetDirectoryName(Document.CurrentPath), Version.ID, Me.FileName)

          Return lstrCurrentPath

        Else
          Return String.Empty
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Friend Sub SetVersion(ByVal lpVersion As Version)
      Try
        mobjVersion = lpVersion
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IMetaHolder Implementation"

    ''' <summary>
    ''' Gets or sets the fully qualified file name for the content element.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property Identifier() As String Implements IMetaHolder.Identifier
      Get
        Return ContentPath
      End Get
      Set(ByVal value As String)
        ContentPath = value
      End Set
    End Property

#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
          Me.Data?.Dispose()
          If Not String.IsNullOrEmpty(Me.ContentPath) Then
            Me.ContentPath = Nothing
          End If
          If Not String.IsNullOrEmpty(Me.FileName) Then
            Me.FileName = Nothing
          End If
          If mobjFileSize IsNot Nothing Then
            mobjFileSize = Nothing
          End If
          mobjMetadata.Dispose()
          mobjProperties.Dispose()
          mobjDocument = Nothing
        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

  End Class

End Namespace