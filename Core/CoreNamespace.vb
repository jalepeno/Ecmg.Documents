'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Security

#If CTSDOTNET = 1 Then
Imports Ecmg.Cts.Arguments
#End If

#If Not SILVERLIGHT = 1 Then
'.Encryption
'Imports System.Windows.Forms

' For Compression Class
'Imports System.IO.Compression
#End If


#End Region

#If Not SILVERLIGHT = 1 Then
''' <summary>
''' Contains the core objects of the Cts framework
''' </summary>
#End If
Namespace Core

  '#If Not SILVERLIGHT = 1 Then
  '  Public Delegate Sub ContentSourceCreatedEventHandler(ByVal sender As Object, ByVal e As Arguments.ContentSourceEventArgs)
  '#End If

#Region "Public Interfaces"

  ''' <summary>
  ''' Implemented by objects that contain metadata
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IMetaHolder

    ''' <summary>
    ''' Gets or sets the unique identifier for the object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Identifier() As String

    ''' <summary>
    ''' Gets or sets the collection of properties for the object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Metadata() As IProperties

  End Interface

  Public Interface IContentPropertyAccessor
    Inherits IMetaHolder
    Inherits IDisposable

    Sub Open(lpFilePath As String)
    Sub Close()
    Sub Save()
    Sub SyncronizeContentProperties(lpSource As IMetaHolder)

    ReadOnly Property IsReadOnly As Boolean

  End Interface

  ''' <summary>
  ''' Imp-lemented by objects that contain auditable events.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface IAuditableItem

    ''' <summary>
    ''' Gets or sets the collection of audit events for the object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property AuditEvents As AuditEvents

  End Interface

  Public Interface IDescription
    Inherits INamedItem

    ''' <summary>
    ''' Gets or sets the name of the item.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Xml.Serialization.XmlAttribute()>
    Overloads Property Name() As String


    '<Xml.Serialization.XmlAttribute()> _
    ''' <summary>
    ''' Gets or sets the description of the item.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property Description() As String

  End Interface

  Public Interface IDisplayable
    Inherits IDescription

    ''' <summary>
    ''' Gets or sets the display name of the item.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Property DisplayName As String

  End Interface

  ''' <summary>
  ''' Used by any object with a Name property
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface INamedItem
    Property Name As String
  End Interface

  ''' <summary>
  ''' Used by collections that can be accessed by by item name.
  ''' </summary>
  ''' <remarks></remarks>
  Public Interface INamedItems
    Sub Add(ByVal item As Object)
    Function Remove(ByVal lpItemName As String) As Boolean
    Property ItemByName(ByVal name As String) As Object
  End Interface

  Public Interface IRepositoryObject
    Inherits INamedItem
    Property Id As String
    Property DescriptiveText As String
    Property DisplayName As String
    ReadOnly Property Permissions As IPermissions

    ''' <summary>
    ''' Gets the set of permissions for this object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If the source provider does not have a 
    ''' security implementation, this list will be empty.
    ''' </remarks>
    ReadOnly Property Properties As IProperties
  End Interface

  ''' <summary>
  ''' Used by objects that can 
  ''' save to or open from an archive.
  ''' </summary>
  ''' <typeparam name="T">
  ''' The type of object to save to 
  ''' or open from the archive.
  ''' </typeparam>
  ''' <remarks>
  ''' Especially useful for hierarchical objects 
  ''' such as Document or Repository which are 
  ''' typically saved to multiple files.
  ''' </remarks>
  Public Interface IArchivable(Of T)
    ''' <summary>
    ''' Gets the default file extension 
    ''' to be used for serialization and 
    ''' deserialization of the archive.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ReadOnly Property DefaultArchiveFileExtension() As String

    Sub Archive(ByVal archivePath As String)
    Sub Archive(ByVal archivePath As String, ByVal removeContainedFiles As Boolean)
    ''' <summary>
    ''' Archives the object
    ''' </summary>
    ''' <param name="archivePath">
    ''' The path to write the archive file to.
    ''' </param>
    ''' <param name="password">
    ''' If desired, the password to secure the archive file.
    ''' </param>
    ''' <param name="removeContainedFiles">
    ''' Specifies whether or not the contained files should 
    ''' be deleted after the archive is created successfully.
    ''' </param>
    ''' <remarks></remarks>
    Sub Archive(ByVal archivePath As String, ByVal removeContainedFiles As Boolean, ByVal password As String)
    Function ToArchiveStream() As IO.Stream
    Function ToArchiveStream(ByVal password As String) As IO.Stream
    Function FromArchive(ByVal archivePath As String) As T
    Function FromArchive(ByVal archivePath As String, ByVal password As String) As T
    Function FromArchive(ByVal stream As IO.Stream) As T
    Function FromArchive(ByVal stream As IO.Stream, ByVal password As String) As T
  End Interface

#End Region

#Region "Public Enumerations"

  Public Enum FileSourceType
    DocumentContent
    File
  End Enum

  ''' <summary>
  ''' Used as a parameter in methods which need to recurse folder structures.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum RecursionLevel
    ecmThisLevelOnly = 1
    ecmThisLevelAndImmediateChildrenOnly = 2
    ecmAllChildren = 3
  End Enum

  ''' <summary>
  ''' Defines the type of object involved in the relationship
  ''' </summary>
  ''' <remarks>Used in the definition of a parent/child document relationship</remarks>
  Public Enum RelatedObjectType
    Document = 0
    Version = 1
  End Enum

  ''' <summary>
  ''' Defines hierarchy in a parent/child document relationship
  ''' </summary>
  ''' <remarks>Used in the definition of a parent/child document relationship</remarks>
  Public Enum RelationshipType
    Child = 1
    Parent = 2
  End Enum

  ''' <summary>
  ''' Defines the strength of a document relationship
  ''' </summary>
  ''' <remarks>Used in the definition of a parent/child document relationship</remarks>
  Public Enum RelationshipStrength
    Weak = 1
    Strong = 2
  End Enum

  ''' <summary>
  ''' Defines the spanning nature of a document relationship
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum RelationshipPersistance
    [Static] = 0
    Dynamic = 1
    DynamicLabel = 2
    URI = 3
  End Enum

  Public Enum RelationshipCascadeDeleteAction
    NoCascadeDelete = 0
    CascadeDelete = 1
  End Enum

  Public Enum RelationshipPreventDeleteAction
    AllowBothDelete = 0
    PreventChildDelete = 1
    PreventParentDelete = 2
    PreventBothDelete = 3
  End Enum

  ''' <summary>
  ''' Define the mode in which a document is to be filed in a folder
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum FilingMode
    UnFiled = 0
    BaseFolderPathOnly = 1
    DocumentFolderPath = 2
    BaseFolderPathPlusDocumentFolderPath = 3
    DocumentFolderPathPlusBaseFolderPath = 4
  End Enum

  ''' <summary>
  ''' Defines how multiple criteria should be evaluated relative to each other for building searches.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum SetEvaluation
    seAnd = 0
    seOr = 1
  End Enum

  ''' <summary>
  ''' Used in part to control which 
  ''' versions of a document are to 
  ''' be exported.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum VersionScopeEnum

    AllVersions = 0
    MostCurrentVersion = 1
    CurrentReleasedVersion = 2
    FirstVersion = 3
    LastNVersions = 4
    FirstNVersions = 5

  End Enum

  Public Enum VersionBindType
    LatestVersion = 0
    LatestMajorVersion = 1
  End Enum

  ''' <summary>
  ''' Used to control whether new items are created as major or minor versions.
  ''' </summary>
  ''' <remarks>Some repositories do not support major vs minor versions.  
  ''' For those repositories this will not apply.</remarks>
  Public Enum VersionTypeEnum
    Unspecified = 0
    Major = 1
    Minor = 2
  End Enum

  ''' <summary>
  ''' Designates a specific Ecmg file format for serialization.
  ''' </summary>
  ''' <remarks>Typically used in the construction of output file names.</remarks>
  Public Enum FileType
    MigrationLog = 0
    MigrationConfigurationFile = 1
    CondensedMigrationConfiguration = 2
    ValidationResultSetOutputFile = 3
    ContentDefinitionFile = 4
    ContentTransformationFile = 5
  End Enum

  Public Enum Result

    ''' <summary>
    ''' Indicates that the operation succeeded.
    ''' </summary>
    ''' <remarks></remarks>
    Success = -1

    ''' <summary>
    ''' Indicates that the operation failed.
    ''' </summary>
    ''' <remarks></remarks>
    Failed = 0

    ''' <summary>
    ''' Indicates the operation has not yet started.
    ''' </summary>
    ''' <remarks></remarks>
    NotProcessed = 1

    ''' <summary>
    ''' Indicates that the operation already succeeded and did not get reprocessed.
    ''' </summary>
    ''' <remarks></remarks>
    PreviouslySucceeded = -101

    ''' <summary>
    ''' Indicates that the operation already failed and did not get reprocessed.
    ''' </summary>
    ''' <remarks></remarks>
    PreviouslyFailed = -100

    ''' <summary>
    ''' Indicates that the operation could not be rolled back 
    ''' because rollback is not supported for the operation.    
    ''' </summary>
    ''' <remarks></remarks>
    RollbackNotSupported = -2

    ''' <summary>
    ''' Indicates that the rollback succeeded.
    ''' </summary>
    ''' <remarks></remarks>
    RollbackSuccess = -11

    ''' <summary>
    ''' Indicates that the rollback failed.
    ''' </summary>
    ''' <remarks></remarks>
    RollbackFailed = -10

    ''' <summary>
    ''' Indicates that the operation could not be rolled back 
    ''' because no rollback implementation is available.    
    ''' </summary>
    ''' <remarks></remarks>
    RollbackNotImplemented = -12

  End Enum

  ''' <summary>Deprecated</summary>
  Public Enum ReplaceResult
    ObjectReplacedPreservingOriginalID = 0
    ObjectReplacedWithNewID = 1
    ObjectNotFoundNoAction = -1
    ObjectNotFoundNewObjectCreated = -2
  End Enum

#End Region

End Namespace