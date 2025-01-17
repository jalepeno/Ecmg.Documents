'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Configuration
Imports System.IO
Imports System.Security
Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Annotations
Imports Documents.Exceptions
Imports Documents.Security
Imports Documents.SerializationUtilities
Imports Documents.Utilities
Imports Ionic.Zip
Imports Newtonsoft.Json
Imports Exception = System.Exception

#End Region

Namespace Core
  ''' <summary>
  ''' Provides an abstract description of a document including all versions with both
  ''' content and metadata.
  ''' </summary>
  ''' <remarks>Serialized instances should use the file extension .cdf (Content Definition File)</remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class Document
    Implements IMetaHolder
    Implements IAuditableItem
    Implements IArchivable(Of Document)

#Region "Class Constants"

    ''' <summary>
    ''' Constant integer flag used in some operations to specify the scope covering all
    ''' versions.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Const ALL_VERSIONS As Integer = -1

    ''' <summary>
    ''' Constant value representing the file extension to save XML serialized instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Const CONTENT_DEFINITION_FILE_EXTENSION As String = "cdf"

    ''' <summary>
    ''' Constant value representing the file extension to save JSON serialized instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public _
      Const JSON_CONTENT_DEFINITION_FILE_EXTENSION As String = "jcdf"

    ''' <summary>
    ''' Constant value representing the file extension to save archived instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Const CONTENT_PACKAGE_FILE_EXTENSION As String = "cpf"

    ''' <summary>
    ''' Constant value representing the file extension to save archived json instances
    ''' to.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public _
      Const JSON_CONTENT_PACKAGE_FILE_EXTENSION As String = "jcpf"

    ''' <summary>
    ''' Constant value specifying the property name used for storing document folder
    ''' paths.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public _
      Const FOLDER_PATHS_PROPERTY_NAME As String = "Folders Filed In"

    ''' <summary>
    ''' Constant value specifying the property name used for storing document folder
    ''' path.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Const FOLDER_PATH_PROPERTY_NAME As String = "Folder Path"

    ''' <summary>
    ''' Alternate constant value specifying the property name used for storing document
    ''' folder path.
    ''' </summary>
    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Public Const FOLDERPATH_PROPERTY_NAME As String = "FolderPath"

#End Region

#Region "Class Variables"

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjID As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mstrHeaderString As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjRelatedDocuments As New Documents

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjRelationships As New Relationships

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjVersions As New Versions(Me)

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjAuditEvents As New AuditEvents

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjProperties As New ECMProperties

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjPermissions As IPermissions = New Permissions

    'Private mstrXMLProcessingInstructions() As String = _
    '          {"xml-stylesheet^type=""text/xsl"" href=""http://Ecmg.us/Ecmg/ecmdocument.xslt""", _
    '           "altova_sps^http://Ecmg.us/Ecmg/ecmdocument.sps"}

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mstrSerializationPath As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mstrDeSerializationPath As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mstrCurrentPath As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mstrCTSVersion As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mstrOriginalLocale As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private _
      menuStorageType As Content.StorageTypeEnum = Content.StorageTypeEnum.Reference

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mstrHash As String = String.Empty

    <DebuggerBrowsable(DebuggerBrowsableState.Never)> Private mobjFolderPathsProperty As ECMProperty

    ' <Added by: Ernie at: 1/11/2013-3:14:43 PM on machine: ERNIE-THINK>
    Private mobjContentDictionary As New Dictionary(Of String, Stream)
    ' </Added by: Ernie at: 1/11/2013-3:14:43 PM on machine: ERNIE-THINK>


    Private mobjTempFilePaths As New List(Of String)
    Private mobjTempFileStreams As New List(Of FileStream)

    'Private mobjTempDocumentFileStream As FileStream
    'Private mstrTempDocumentFilePath As String

    'Private mobjTempContentFileStream As FileStream
    'Private mstrTempContentFilePath As String

#End Region

#Region "Public Properties"

#Region "IMetaHolder Implementation"

    ''' <summary>
    ''' Gets or sets the unique ID of the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This is the same as the ID property.</remarks>
    <JsonIgnore()>
    Public Property Identifier() As String Implements IMetaHolder.Identifier
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return ID
      End Get
      Set(ByVal value As String)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ID = value
      End Set
    End Property

#End Region

    ''' <summary>
    ''' The ID of the document
    ''' </summary>
    <XmlAttribute()>
    Public Property ID() As String
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        If mobjID Is Nothing Then
          Return Me.Name
        End If
        Return mobjID
      End Get
      Set(ByVal value As String)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mobjID = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the class of document to which this document belongs.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DocumentClass() As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If Properties.PropertyExists("Document Class") = False Then
            'Properties.Add(New ECMProperty(PropertyType.ecmString, _
            '                               "Document Class", ""))
            Properties.Add(PropertyFactory.Create(PropertyType.ecmString,
                                                  "Document Class", String.Empty))
          End If
          Return Properties("Document Class").Value.ToString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return String.Empty
        End Try
      End Get

      Set(ByVal Value As String)
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If Properties.PropertyExists("Document Class") Then
            Properties("Document Class").Value = Value
          Else
            'Properties.Add(New ECMProperty(PropertyType.ecmString, _
            '                               "Document Class", Value))
            Properties.Add(PropertyFactory.Create(PropertyType.ecmString,
                                                  "Document Class", Value))

          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Document::Set_DocumentClass('{0}')", Value))
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The name of the locale used in the 
    ''' original creation or export of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute("OriginalLocale")>
    Public Property OriginalLocale() As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If mstrOriginalLocale.Length = 0 Then
            mstrOriginalLocale = Globalization.CultureInfo.CurrentCulture.Name
          End If
          Return mstrOriginalLocale
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return String.Empty
        End Try
      End Get
      Set(ByVal value As String)
        mstrOriginalLocale = value
      End Set
    End Property

    <XmlAttribute("Summary")>
    Public Property Summary As String
      Get
        Try
          Return DebuggerIdentifier()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          If Helper.IsDeserializationBasedCall OrElse
             Helper.CallStackContainsMethodName("AssignObjectProperties") OrElse
             Helper.CallStackContainsMethodName("Clone") Then
            ' Just ignore since we don't store this anyway
            ' It is just here for writing out to the xml
          Else
            Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets the total content size of all versions.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TotalContentSize As FileSize
      Get
        Try
          Return GetTotalContentSize()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property


    ''' <summary>
    ''' The version of Content Transformation Services used in the 
    ''' original creation or export of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute("CtsVersion")>
    <JsonProperty("ctsVersion")>
    Public Property CTS_Version() As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If mstrCTSVersion.Length = 0 Then
            mstrCTSVersion = FrameworkVersion.CurrentVersion
          End If

          Return mstrCTSVersion

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return String.Empty
        End Try
      End Get
      Set(ByVal value As String)
        mstrCTSVersion = value
      End Set
    End Property


    ''' <summary>
    ''' Contains the encrypted document header
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Although this may appear to be a read/write property, 
    ''' any attempt to set the property value will result in an InvalidOperationException.  
    ''' The property set operation is reserved for internal use only.</remarks>
    <XmlAttribute("Header")>
    <JsonIgnore()>
    Public Property HeaderString() As String
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          Return mstrHeaderString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If Helper.IsDeserializationBasedCall OrElse
             Helper.CallStackContainsMethodName("AssignObjectProperties") OrElse
             Helper.CallStackContainsMethodName("Clone") Then
            mstrHeaderString = value
            ''  Recreate the header
            'UpdateHeader(value)
          Else
            Throw New InvalidOperationException(TREAT_AS_READ_ONLY)
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The object identifier for the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>This value is typically initialized to the new document 
    ''' identifier during a document add or import operation.</remarks>
    <JsonIgnore()>
    Public Property ObjectID() As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If Properties.PropertyExists("ObjectID") = False Then
            'Properties.Add(New ECMProperty(PropertyType.ecmString, _
            '                               "ObjectID", ""))
            Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "ObjectID", String.Empty))
          End If
          Return Properties("ObjectID").Value.ToString
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return String.Empty
        End Try
      End Get
      Set(ByVal Value As String)
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          If Properties.PropertyExists("ObjectID") Then
            Properties("ObjectID").Value = Value
          Else
            'Properties.Add(New ECMProperty(PropertyType.ecmString, _
            '                               "ObjectID", Value))
            Properties.Add(PropertyFactory.Create(PropertyType.ecmString, "ObjectID", Value))
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, String.Format("Value::Set_ObjectID('{0}')", Value))
        End Try
      End Set
    End Property

#Region "IAuditableItem Implementation"

    Public Property AuditEvents As AuditEvents Implements IAuditableItem.AuditEvents
      Get
        Return mobjAuditEvents
      End Get
      Set(value As AuditEvents)
        mobjAuditEvents = value
      End Set
    End Property

#End Region

    ''' <summary>
    ''' Gets or sets the parent and/or child relationships for the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Relationships() As Relationships
      Get
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          Return mobjRelationships

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Relationships)
        Try
#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          mobjRelationships = value

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the collection of properties for the document.
    ''' </summary>
    Public Property Properties() As ECMProperties
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return mobjProperties
      End Get
      Set(ByVal value As ECMProperties)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mobjProperties = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the permissions for the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Permissions As Permissions
      Get
        Try
          Return mobjPermissions
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Permissions)
        Try
          mobjPermissions = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The name of the document.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If no property is defined for 'Name' 
    ''' this may also return the value of the following 
    ''' properties respectively, 
    ''' ('Title', 'DocumentTitle', 'Document Title', 'FileName', 'Subject')</remarks>
    Public ReadOnly Property Name() As String
      Get
        Try

#If NET8_0_OR_GREATER Then
          ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
          Dim lobjNameValue As Object
          If Me.PropertyExists(PropertyScope.DocumentProperty, "Name") Then
            lobjNameValue = Properties("Name").Value
          ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "Title") Then
            lobjNameValue = GetProperty("Title", PropertyScope.VersionProperty).Value
          ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "DocumentTitle") Then
            lobjNameValue = GetProperty("DocumentTitle", PropertyScope.VersionProperty).Value _
            'Properties("DocumentTitle").Value
          ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "Document Title") Then
            lobjNameValue = GetProperty("Document Title", PropertyScope.VersionProperty).Value _
            'Properties("DocumentTitle").Value
          ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "FileName") Then
            lobjNameValue = GetProperty("FileName", PropertyScope.VersionProperty).Value 'Properties("Subject").Value
          ElseIf Me.PropertyExists(PropertyScope.VersionProperty, "Subject") Then
            lobjNameValue = GetProperty("Subject", PropertyScope.VersionProperty).Value 'Properties("Subject").Value

          ElseIf mobjID IsNot Nothing Then
            lobjNameValue = New Value(ID)
          Else
            lobjNameValue = New Value("")
          End If

          'If (lobjNameValue.GetType().IsInstanceOfType(GetType(Value))) Then
          '  Return lobjNameValue.Value.ToString
          'ElseIf lobjNameValue.ToString = "Ecmg.Cts.Core.Value" Then
          '  Return lobjNameValue.value
          'Else
          '  Return lobjNameValue.ToString()
          'End If

          If TypeOf lobjNameValue Is Value Then
            Return lobjNameValue.Value.ToString
          ElseIf lobjNameValue.ToString.EndsWith("Core.Value") Then
            Return lobjNameValue.value
          Else
            Return lobjNameValue.ToString()
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return ""
        End Try
      End Get
    End Property

    ''' <summary>
    ''' The path the document was de-serialized from
    ''' </summary>
    ''' <value></value>
    ''' <returns>The fully qualifed cdf path.</returns>
    ''' <remarks>Only present when the document has been de-serialized from a cdf file.</remarks>
    Public ReadOnly Property CurrentPath() As String
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return mstrCurrentPath
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the versions of the document.
    ''' </summary>
    Public Property Versions() As Versions
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return mobjVersions
      End Get
      Set(ByVal value As Versions)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mobjVersions = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the collection of metadata properties for the object.
    ''' </summary>
    ''' <returns>ECMProperties collection</returns>
    ''' <remarks>Interface passthrough to the Properties collection.</remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property Metadata() As IProperties Implements IMetaHolder.Metadata
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return Properties
      End Get
      Set(ByVal value As IProperties)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Properties = value
      End Set
    End Property

    ''' <summary>
    ''' The path the document was serialized to.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property SerializationPath() As String
      ' Intended to be used for re-writing out file after transformation
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return mstrSerializationPath
      End Get
      Set(ByVal value As String)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mstrSerializationPath = value
      End Set
    End Property

    ''' <summary>
    ''' The path the document was deserialized to.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    <JsonIgnore()>
    Public Property DeSerializationPath() As String
      ' Intended to be used for re-writing out file after transformation
      Get

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return mstrDeSerializationPath
      End Get
      Set(ByVal value As String)

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        mstrDeSerializationPath = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the type of storage associated with the document content
    ''' </summary>
    ''' <value>The StorageTypeEnum value to set for the Document StorageType</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public Property StorageType() As Content.StorageTypeEnum
      Get


#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        Return menuStorageType
      End Get

      Set(ByVal value As Content.StorageTypeEnum)


#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If
        menuStorageType = value

        For Each lobjVersion As Version In Versions
          For Each lobjContent As Content In lobjVersion.Contents
            lobjContent.StorageType = value
          Next
        Next
      End Set
    End Property

    ''' <summary>
    ''' Gets the first version of the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property FirstVersion As Version
      Get
        Try
          Return GetFirstVersion()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets the latest version of the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <JsonIgnore()>
    Public ReadOnly Property LatestVersion As Version
      Get
        Try
          Return GetLatestVersion()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Friend Properties"

    Friend ReadOnly Property TempFilePaths As List(Of String)
      Get
        Try
          Return mobjTempFilePaths
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Friend ReadOnly Property TempFileStreams As List(Of FileStream)
      Get
        Try
          Return mobjTempFileStreams
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpByteArray As Byte())
      Try

        Dim lobjMemoryStream As New IO.MemoryStream(lpByteArray)
        Dim lobjDocument As New Document(lobjMemoryStream)

        Helper.AssignObjectProperties(lobjDocument, Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Constructs a new Document object from the specified stream.
    ''' </summary>
    ''' <param name="lpStream">The stream to create the document from.</param>
    ''' <remarks>Will attempt to create the document from 
    ''' either a native stream or an archive stream.</remarks>
    Public Sub New(ByVal lpStream As Stream)
      Try

        Dim lobjDocument As Document = Nothing

        Dim lobjSeekableStream As Stream = lpStream

        ' If we cannot seek, then create a copy of the stream
        ' in memory and pass that along
        If (lobjSeekableStream.CanSeek = False) Then

          lobjSeekableStream = New MemoryStream()
          Dim buf(1023) As Byte
          Dim numRead As Integer = lpStream.Read(buf, 0, 1024)
          While (numRead > 0)
            lobjSeekableStream.Write(buf, 0, numRead)
            numRead = lpStream.Read(buf, 0, 1024)
          End While

          lobjSeekableStream.Position = 0

        End If

#If SILVERLIGHT <> 1 Then
        ' Check to see if the stream is for an xml file or not
        If Helper.IsXmlStream(lobjSeekableStream) Then
          ' The stream is a native xml file, this would be for a cdf file.
          lobjDocument = CreateFromStream(lobjSeekableStream)
        ElseIf Helper.IsZipStream(lobjSeekableStream) Then
          ' The stream is not a native xml file, this would most likely be for a cpf file.
          lobjDocument = FromArchive(lobjSeekableStream)
        Else
          Throw New DocumentException(String.Empty, "An xml stream or a zip stream was expected.")
        End If
#Else
        lobjDocument = FromArchive(lobjSeekableStream)
#End If

        Helper.AssignObjectProperties(lobjDocument, Me)

        SetDocumentReferences()

#If SILVERLIGHT <> 1 Then
        ' We are checking to see if the existing 'cdf' file has a header
        ' If it does not have a header we want to skip over the update header
        If lobjDocument.HeaderString.Length > 0 Then
          '  Recreate the header
          UpdateHeader()
        End If
#End If

        mstrCurrentPath = lobjDocument.CurrentPath

      Catch BadPasswordEx As BadPasswordException
        Dim _
          lobjDocumentEx As _
            New DocumentException(String.Empty, "A password was expected but not supplied.", BadPasswordEx)
        ApplicationLogging.LogException(lobjDocumentEx, Reflection.MethodBase.GetCurrentMethod)
        Throw lobjDocumentEx
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Constructs a new Document object from the specified archive stream and password.
    ''' </summary>
    ''' <param name="lpStream">The archive stream to create the document from.</param>
    ''' <param name="lpPassword">The password to unlock the archive stream.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpStream As Stream, ByVal lpPassword As String)
      Try

        Dim lobjDocument As Document = FromArchive(lpStream, lpPassword)

        Helper.AssignObjectProperties(lobjDocument, Me)

        SetDocumentReferences()

#If SILVERLIGHT <> 1 Then
        ' We are checking to see if the existing 'cdf' file has a header
        ' If it does not have a header we want to skip over the update header
        If lobjDocument.HeaderString.Length > 0 Then
          '  Recreate the header
          UpdateHeader()
        End If
#End If

        mstrCurrentPath = lobjDocument.CurrentPath

      Catch BadPasswordEx As BadPasswordException
        Dim lobjDocumentEx As New DocumentException(String.Empty, "The specified password does not match.")
        ApplicationLogging.LogException(lobjDocumentEx, Reflection.MethodBase.GetCurrentMethod)
        Throw lobjDocumentEx
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Properties"

    Public ReadOnly Property RelatedDocuments As Documents
      Get
        Return mobjRelatedDocuments
      End Get
    End Property

#End Region

#Region "Friend Properties"

    ' <Added by: Ernie at: 1/11/2013-3:21:37 PM on machine: ERNIE-THINK>
    ' ''' <summary>
    ' '''     Used to manage a unique set of content for the document.
    ' ''' </summary>
    ' ''' <value>
    ' '''     <para>
    ' '''         
    ' '''     </para>
    ' ''' </value>
    ' ''' <remarks>
    ' '''     This is intended to support content de-duplication within a single document object.
    ' ''' </remarks>
    'Friend ReadOnly Property ContentDictionary As Dictionary(Of String, Stream)
    '  Get
    '    Try
    '      Return mobjContentDictionary
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    'End Property
    ' </Added by: Ernie at: 1/11/2013-3:21:37 PM on machine: ERNIE-THINK>

#End Region

#Region "Public Methods"

    Public Function GetAllPropertyDefinitions() As IProperties
      Try
        Dim lobjProperties As New ECMProperties
        For Each lobjProperty As IProperty In Me.Properties
          If lobjProperties.PropertyExists(lobjProperty.Name) = False Then
            lobjProperties.Add(lobjProperty)
          End If
        Next

        For Each lobjVersion As Version In Me.Versions
          For Each lobjProperty As IProperty In lobjVersion.Properties
            If lobjProperties.PropertyExists(lobjProperty.Name) = False Then
              lobjProperties.Add(lobjProperty)
            End If
          Next
        Next

        lobjProperties.Sort()

        Return lobjProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Gets all the properties in the collection of the specified type.
    ''' </summary>
    ''' <param name="lpPropertyType">The property type to select</param>
    ''' <returns></returns>
    ''' <remarks>If no properties exist in the collection for the specified 
    ''' type, an empty collection is returned.</remarks>
    Public Function PropertiesByType(lpPropertyType As PropertyType) As ECMProperties
      Try
        Return mobjProperties.PropertiesByType(lpPropertyType)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Determine if a property exists in the document.
    ''' </summary>
    ''' <param name="lpName">The name of the property to check</param>
    ''' <returns></returns>
    ''' <remarks>Checks in both document and version properties</remarks>
    Public Function PropertyExists(ByVal lpName As String, ByVal lpCaseSensitive As Boolean) As Boolean
      Try
#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return PropertyExists(PropertyScope.BothDocumentAndVersionProperties, lpName, lpCaseSensitive)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Determine if a property exists in the document.
    ''' </summary>
    ''' <param name="lpName">The name of the property to check</param>
    ''' <returns></returns>
    ''' <remarks>Checks in both document and version properties</remarks>
    Public Function PropertyExists(ByVal lpName As String) As Boolean
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return PropertyExists(PropertyScope.BothDocumentAndVersionProperties, lpName, True)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="lpPropertyScope"></param>
    ''' <param name="lpName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PropertyExists(ByVal lpPropertyScope As PropertyScope, ByVal lpName As String) As Boolean

#If NET8_0_OR_GREATER Then
      ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

      Return PropertyExists(lpPropertyScope, lpName, True)
    End Function

    ''' <summary>
    ''' Determine if a property exists in the document.
    ''' </summary>
    ''' <param name="lpName">The name of the property to check</param>
    ''' <param name="lpPropertyScope">Specifies whether to check document scoped properties, version scoped properties or both</param>
    ''' <remarks></remarks>
    Public Function PropertyExists(ByVal lpPropertyScope As PropertyScope, ByVal lpName As String,
                                   ByVal lpCaseSensitive As Boolean) As Boolean

      ' ApplicationLogging.WriteLogEntry("Enter Document::PropertyExists", TraceEventType.Verbose)
      Dim lblnPropertyExists As Boolean = False

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Select Case lpPropertyScope
          Case PropertyScope.DocumentProperty
            lblnPropertyExists = Properties.PropertyExists(lpName, lpCaseSensitive)

          Case PropertyScope.VersionProperty
            For Each lobjVersion As Version In Versions
              lblnPropertyExists = lobjVersion.Properties.PropertyExists(lpName, lpCaseSensitive)
              If lblnPropertyExists = True Then
                Exit For
              End If
            Next

          Case PropertyScope.BothDocumentAndVersionProperties
            lblnPropertyExists = Properties.PropertyExists(lpName, lpCaseSensitive)
            If lblnPropertyExists = True Then
              ' If we already found the property then we can exit here
              'Return lblnPropertyExists
              Exit Select
            End If

            For Each lobjVersion As Version In Versions
              lblnPropertyExists = lobjVersion.Properties.PropertyExists(lpName, lpCaseSensitive)
              If lblnPropertyExists = True Then
                Exit For
              End If
            Next
            Exit Select

          Case Else
            Exit Select

        End Select

        Return lblnPropertyExists

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        'ApplicationLogging.WriteLogEntry("Exit Document::PropertyExists", TraceEventType.Verbose)
        Return lblnPropertyExists
      Finally
        'ApplicationLogging.WriteLogEntry("Exit Document::PropertyExists", TraceEventType.Verbose)
      End Try
    End Function

    ''' <summary>
    ''' Gets a specific property by name
    ''' </summary>
    ''' <param name="lpPropertyName"></param>
    ''' <param name="lpVersionIndex"></param>
    ''' <returns></returns>
    ''' <remarks>Checks for the property first at the docuemnt level, then at the version level.</remarks>
    Public Function GetProperty(ByVal lpPropertyName As String,
                                Optional ByVal lpVersionIndex As Integer = 0) As ECMProperty
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Return GetProperty(lpPropertyName, PropertyScope.BothDocumentAndVersionProperties, lpVersionIndex)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Gets a specific property by name
    ''' </summary>
    ''' <param name="lpPropertyName"></param>
    ''' <param name="lpPropertyScope"></param>
    ''' <param name="lpVersionIndex"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetProperty(ByVal lpPropertyName As String,
                                ByVal lpPropertyScope As PropertyScope,
                                Optional ByVal lpVersionIndex As Integer = 0) As ECMProperty

      'ApplicationLogging.WriteLogEntry("Enter Document::GetProperty", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Select Case lpPropertyScope
          Case PropertyScope.DocumentProperty
            Return Properties(lpPropertyName)

          Case PropertyScope.VersionProperty
            If lpVersionIndex = ALL_VERSIONS Then
              Return GetLatestVersion.Properties(lpPropertyName)
            Else
              Return Versions(lpVersionIndex).Properties(lpPropertyName)
            End If

          Case PropertyScope.BothDocumentAndVersionProperties
            If Properties.PropertyExists(lpPropertyName) Then
              Return Properties(lpPropertyName)
            ElseIf Versions(lpVersionIndex).Properties.PropertyExists(lpPropertyName) Then
              Return Versions(lpVersionIndex).Properties(lpPropertyName)
            Else
              Throw _
                New Exceptions.PropertyDoesNotExistException(
                  "Property '" & lpPropertyName & "' does not exist in the document.", lpPropertyName)
            End If

          Case PropertyScope.ContentProperty
            If lpVersionIndex = ALL_VERSIONS Then
              'Return GetLatestVersion.GetPrimaryContent.Metadata.Item(lpPropertyName)
              Return GetLatestVersion.GetPrimaryContent.CompleteProperties(lpPropertyName)
            Else
              'Return Versions(lpVersionIndex).GetPrimaryContent.Metadata.Item(lpPropertyName)
              Return Versions(lpVersionIndex).GetPrimaryContent.CompleteProperties(lpPropertyName)
            End If

          Case PropertyScope.AllProperties
            If Properties.PropertyExists(lpPropertyName) Then
              Return Properties(lpPropertyName)
            ElseIf Versions(lpVersionIndex).Properties.PropertyExists(lpPropertyName) Then
              Return Versions(lpVersionIndex).Properties(lpPropertyName)
            ElseIf _
              GetLatestVersion.HasContent = True AndAlso
              GetLatestVersion.GetPrimaryContent.Metadata.PropertyExists(lpPropertyName, False) Then
              Return GetLatestVersion.GetPrimaryContent.Metadata.Item(lpPropertyName)
            Else
              Throw _
                New Exceptions.PropertyDoesNotExistException(
                  "Property '" & lpPropertyName & "' does not exist in the document.", lpPropertyName)
            End If

          Case Else
            Return Properties(lpPropertyName)

        End Select
      Catch propEx As PropertyDoesNotExistException
        ApplicationLogging.LogException(propEx, Reflection.MethodBase.GetCurrentMethod)
        Throw
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Throw New InvalidOperationException("Unable to Get Property '" & lpPropertyName & "'", ex)
      Finally
        'ApplicationLogging.WriteLogEntry("Exit Document::GetProperty", TraceEventType.Verbose)
      End Try
    End Function

    ''' <summary>
    ''' Removes all unsettable properties from the document
    ''' </summary>
    ''' <param name="lpDocumentClass"></param>
    ''' <remarks></remarks>
    Public Sub RemoveUnsettableProperties(ByVal lpDocumentClass As DocumentClass)
      Try

        Dim lobjClassificationProperty As ClassificationProperty = Nothing
        Dim lstrPropertyName As String = String.Empty

        ' Remove any read only properties from the document level
        Properties.RemoveReadOnlyProperties(lpDocumentClass)


        'For lintVersionCounter As Integer = 1 To Versions.Count - 1
        For Each lobjVersion As Version In Versions
          ' Remove any read only properties from the document level
          lobjVersion.Properties.RemoveReadOnlyProperties(lpDocumentClass)

          ' Clear any settable only on create for all but the first version.
          If lobjVersion.ID > 0 Then
            For lintPropertyCounter As Integer = lobjVersion.Properties.Count - 1 To 0 Step -1
              lstrPropertyName = lobjVersion.Properties(lintPropertyCounter).Name
              lobjClassificationProperty = lpDocumentClass.Properties.ItemByName(lstrPropertyName)
              If lobjClassificationProperty IsNot Nothing Then
                If _
                  lobjClassificationProperty.Settability =
                  ClassificationProperty.SettabilityEnum.SETTABLE_ONLY_ON_CREATE Then
                  lobjVersion.Properties.Remove(lobjVersion.Properties(lintPropertyCounter))
                  ApplicationLogging.WriteLogEntry(
                    String.Format("Removed settable only on create property '{0}' from version {1} of document '{2}'",
                                  lobjClassificationProperty.Name, lobjVersion.ID, Name),
                    TraceEventType.Information, 4255)
                End If
              End If
            Next
          End If
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#Region "Document Version Methods"

    Public Function GetFirstVersion() As Version

      ApplicationLogging.WriteLogEntry("Enter Document::GetFirstVersion", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ApplicationLogging.WriteLogEntry("Exit Document::GetFirstVersion", TraceEventType.Verbose)
        If Versions.Count = 0 Then
          Versions.Add(New Version("0"))
        End If

        For Each lobjVersion As Version In Versions
          Return lobjVersion
        Next

        Throw New Exceptions.DocumentException(Me, "There are no versions present in this document.")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ApplicationLogging.WriteLogEntry("Exit Document::GetFirstVersion", TraceEventType.Verbose)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Returns the latest version of the document
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>If the document has no versions defined, 
    ''' this method will add a new empty version to the 
    ''' document and return the new version.</remarks>
    Public Function GetLatestVersion() As Version
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        If Versions.Count = 0 Then
          Versions.Add(New Version("0"))
        End If
        Return Versions(Versions.Count - 1)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

    ''' <summary>
    ''' Encodes all the content data per for current storage type
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub EncodeAllContents()

#If NET8_0_OR_GREATER Then
      ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

      EncodeAllContents(Me.StorageType)
    End Sub

    ''' <summary>
    ''' Sets the current StorageType value and encodes all the content data per the specified storage type
    ''' </summary>
    ''' <param name="lpStorageType">The storage type for the document content</param>
    ''' <remarks></remarks>
    Public Sub EncodeAllContents(ByVal lpStorageType As Content.StorageTypeEnum)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        '  Only proceed if the current storage type does not match the one specified here
        If StorageType = lpStorageType Then
          Exit Sub
        End If

        '  First set the document storage type to match the one specified here
        StorageType = lpStorageType

        '  Set the data accordingly for each content in every version
        For Each lobjVersion As Version In Me.Versions
          For Each lobjContent As Content In lobjVersion.Contents
            lobjContent.SetData(lpStorageType)
          Next
        Next

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Determines whether or not the document has any content.
    ''' </summary>
    ''' <returns>True if content is present, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Function HasContent() As Boolean
      ' ApplicationLogging.WriteLogEntry("Enter Document::HasContent", TraceEventType.Verbose)
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        Dim lintContentCount As Integer = 0
        For Each lobjVersion As Version In Versions
          lintContentCount += lobjVersion.Contents.Count
        Next
        If lintContentCount > 0 Then
          ' ApplicationLogging.WriteLogEntry("Exit Document::HasContent", TraceEventType.Verbose)
          Return True
        Else
          ' ApplicationLogging.WriteLogEntry("Exit Document::HasContent", TraceEventType.Verbose)
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' ApplicationLogging.WriteLogEntry("Exit Document::HasContent", TraceEventType.Verbose)
        Return False
      End Try
    End Function

    ''' <summary>
    ''' Returns the number of invalid relationships.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InvalidRelationshipCount() As Integer
      Try
        If Relationships Is Nothing OrElse Relationships.Count = 0 Then
          Return 0
        Else
          Return Relationships.InvalidRelationshipCount(Me)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Recurses through each Version and Content and sets the parent Document reference
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetDocumentReferences()
      Try
        If Versions IsNot Nothing Then
          Versions.SetDocument(Me)
          For Each lobjVersion As Version In Versions
            lobjVersion.SetDocument(Me)
            For Each lobjContent As Content In lobjVersion.Contents
              lobjContent.SetVersion(lobjVersion)
            Next
          Next
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Creates and returns a stream of the current document.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' If you want the stream to include all of the 
    ''' document content and the storage type is not 
    ''' encoded, use ToArchiveStream instead.
    ''' </remarks>
    Public Function ToStream() As IO.MemoryStream

      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        SetDocumentReferences()

        '  Encode the contents as appropriate for the storage type
        EncodeAllContents()

        Return Serializer.Serialize.ToStream(Me)

        'Return ToArchiveStream()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try

        If IsDisposed Then
          Return "Document Disposed"
        End If

        Dim lstrName As String = Name

        lobjIdentifierBuilder.AppendFormat("{0} - ", ID)

        ' 1st show the class, if assigned.
        If DocumentClass.Length > 0 Then
          lobjIdentifierBuilder.AppendFormat("({0}) ", DocumentClass)
        End If

        If lstrName Is Nothing OrElse lstrName.Length = 0 Then
          lobjIdentifierBuilder.Append("Document identifier not set")
        Else
          lobjIdentifierBuilder.AppendFormat("{0}", lstrName)
        End If

        If Versions Is Nothing OrElse Versions.Count = 0 Then
          lobjIdentifierBuilder.Append(": No Versions")
        ElseIf Versions.Count = 1 Then
          lobjIdentifierBuilder.AppendFormat(": 1 version ({0})", Me.GetFirstVersion.DebuggerIdentifier)
        ElseIf Versions.Count > 1 Then
          lobjIdentifierBuilder.AppendFormat(": {0} versions: contents {1}", Versions.Count,
                                             TotalContentSize.ToString)
        End If

        Dim lintInvalidRelationshipCount As Integer = InvalidRelationshipCount()

        If lintInvalidRelationshipCount > 0 Then
          lobjIdentifierBuilder.AppendFormat(" * {0} Relationship(s), {1} Invalid", Relationships.Count,
                                             InvalidRelationshipCount())
          'ElseIf lintInvalidRelationshipCount = 1 Then
          '  lobjIdentifierBuilder.AppendFormat(" * {0} Relationships, 1 Invalid", Relationships.Count)
        ElseIf Relationships IsNot Nothing AndAlso Relationships.Count > 0 Then
          lobjIdentifierBuilder.AppendFormat(" * {0} Relationship(s)", Relationships.Count)
        End If

        Return lobjIdentifierBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Function GetTotalContentSize() As FileSize
      Try
        Dim llngTotalContentBytes As Long = 0
        If HasContent() Then
          For Each lobjVersion As Version In Versions
            llngTotalContentBytes += lobjVersion.TotalContentSize.Bytes
          Next
        End If
        Return New FileSize(llngTotalContentBytes)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IArchivable(Of Document) Implementation"

    <JsonIgnore()>
    Public ReadOnly Property DefaultArchiveFileExtension() As String _
      Implements IArchivable(Of Document).DefaultArchiveFileExtension
      Get
        Return CONTENT_PACKAGE_FILE_EXTENSION
      End Get
    End Property

    ''' <summary>
    ''' Create a Document object from a cdfa file.
    ''' </summary>
    ''' <param name="archivePath">
    ''' A fully qualified path to a Content 
    ''' Definition File Archive (cdfa) file.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FromArchive(ByVal archivePath As String) As Document Implements IArchivable(Of Document).FromArchive
      Try


        ' Check to see if we have a valid archivePath
        Helper.VerifyFilePath(archivePath, "archivePath", True)

        If ZipFile.IsZipFile(archivePath) = False Then
          Throw New Exceptions.InvalidPathException(
            String.Format("The file '{0}' is not a valid archive file", archivePath), archivePath)
        End If

        'Dim lobjZipStream As MemoryStream = Serializer.Deserialize.ReadFileToMemory(archivePath)
        Dim lobjZipStream As Stream

        Dim lobjFileSize As New FileSize(archivePath)

        If lobjFileSize.Megabytes < ConfigurationManager.AppSettings("MaximumInMemoryDocumentMegabytes") Then
          lobjZipStream = Serializer.Deserialize.ReadFileToMemory(archivePath)
        Else
          lobjZipStream = New FileStream(archivePath, FileMode.Open)
        End If

        Return FromArchive(lobjZipStream)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Create a Document object from a cdfa file.
    ''' </summary>
    ''' <param name="archivePath">
    ''' A fully qualified path to a Content 
    ''' Definition File Archive (cdfa) file.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FromArchive(ByVal archivePath As String, ByVal password As String) As Document _
      Implements IArchivable(Of Document).FromArchive
      Try


        ' Check to see if we have a valid archivePath
        Helper.VerifyFilePath(archivePath, "archivePath", True)

        If ZipFile.IsZipFile(archivePath) = False Then
          Throw New Exceptions.InvalidPathException(
            String.Format("The file '{0}' is not a valid archive file", archivePath), archivePath)
        End If

        Dim lobjZipFile As ZipFile
        lobjZipFile = New ZipFile

        Dim lobjZipStream As MemoryStream = Serializer.Deserialize.ReadFileToMemory(archivePath)

        Return FromArchive(lobjZipStream, password)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Create a Document object from a stream
    ''' </summary>
    ''' <param name="stream">
    ''' A stream containing a DotNetZip ZipFile 
    ''' object with a document object as the first entry.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FromArchive(ByVal stream As System.IO.Stream) As Document _
      Implements IArchivable(Of Document).FromArchive
      Try

        stream.Position = 0
        Dim lobjZipFile As ZipFile = ZipFile.Read(stream)

        Return FromZip(lobjZipFile)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Create a Document object from a stream
    ''' </summary>
    ''' <param name="stream">
    ''' A stream containing a DotNetZip ZipFile 
    ''' object with a document object as the first entry.
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FromArchive(ByVal stream As System.IO.Stream,
                                ByVal password As String) As Document _
      Implements IArchivable(Of Document).FromArchive
      Try

        stream.Position = 0
        Dim lobjZipFile As ZipFile = ZipFile.Read(stream)

        Return FromZip(lobjZipFile, password)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a Document object from a ZipFile object reference
    ''' </summary>
    ''' <param name="zip">A DotNetZip ZipFile object</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Currently coerces the StorageType of the Document 
    ''' to EncodedUnCompressed, need to add a StorageType 
    ''' parameter to determine the desired StorageType or 
    ''' allow it to remain as originally created.
    ''' </remarks>
    Public Function FromZip(ByVal zip As Object) As Document
      Try
        Return FromZip(zip, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a Document object from a ZipFile object reference
    ''' </summary>
    ''' <param name="zip">A DotNetZip ZipFile object</param>
    ''' <param name="password">The password to open the zip file</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Currently coerces the StorageType of the Document 
    ''' to EncodedUnCompressed, need to add a StorageType 
    ''' parameter to determine the desired StorageType or 
    ''' allow it to remain as originally created.
    ''' </remarks>
    Public Function FromZip(ByVal zip As Object,
                            ByVal password As String) As Document
      Try

#If NET8_0_OR_GREATER Then
        ObjectDisposedException.ThrowIf(IsDisposed, Me)
#Else
        If IsDisposed Then
          Throw New ObjectDisposedException(Me.GetType.ToString)
        End If
#End If

        ' Check to see if we have a valid archivePath
        If zip Is Nothing Then
          Throw New ArgumentNullException(NameOf(zip))
        ElseIf TypeOf (zip) Is ZipFile = False Then
          Throw New ArgumentException(
            String.Format("The type '{0}' was not expected, zip must be of type ZipFile.",
                          zip.GetType.Name), NameOf(zip))
        End If

        Dim lobjZipFile As ZipFile = zip

        If lobjZipFile IsNot Nothing Then

          ' If supplied, set the password
          If password IsNot Nothing AndAlso password.Length > 0 Then
            lobjZipFile.Password = password
          End If

          ' Check to make sure the first entry is a cdf file
          'If lobjZipFile.Entries.Count = 0 OrElse _
          '  lobjZipFile.Entries(0).FileName.EndsWith(CONTENT_DEFINITION_FILE_EXTENSION) = False Then
          If lobjZipFile.Entries.Count = 0 OrElse
             lobjZipFile.Item(0).FileName.EndsWith(CONTENT_DEFINITION_FILE_EXTENSION) = False Then
            Throw New Exceptions.DocumentException(String.Empty,
                                                   String.Format("The archive '{0}' does not begin with a CDF Document.",
                                                                 lobjZipFile.Name))
          End If

          Dim lobjDocumentStream As New MemoryStream
          'lobjZipFile.Entries(0).Extract(lobjDocumentStream)
          lobjZipFile.Item(0).Extract(lobjDocumentStream)

          Dim lobjDocument As Document

#If Not SILVERLIGHT = 1 Then
          If lobjZipFile.Item(0).FileName.EndsWith(JSON_CONTENT_DEFINITION_FILE_EXTENSION) Then
            lobjDocument = Document.FromJson(Helper.CopyStreamToString(lobjDocumentStream))

          ElseIf lobjZipFile.Item(0).FileName.EndsWith(CONTENT_DEFINITION_FILE_EXTENSION) Then
            lobjDocument = Serializer.Deserialize.FromStream(lobjDocumentStream, Me.GetType)

          Else
            Throw New DocumentValidationException(lobjZipFile.Item(0).FileName,
                                                  String.Format("{0} or {1} expected.",
                                                                CONTENT_DEFINITION_FILE_EXTENSION,
                                                                JSON_CONTENT_DEFINITION_FILE_EXTENSION))
          End If
#Else
          lobjDocument = Serializer.Deserialize.FromStream(lobjDocumentStream, Me.GetType)
#End If


          ' <Modified by: Ernie at 9/20/2011-8:22:14 AM on machine: ERNIE-M4400>
          ' We are not able to support related documents with Silverlight at this time
#If Not SILVERLIGHT = 1 Then

          ' Get any related documents
          GetRelatedDocuments(lobjDocument, lobjZipFile)

#End If
          ' </Modified by: Ernie at 9/20/2011-8:22:14 AM on machine: ERNIE-M4400>


          ' Extract the content from the document
          ExtractContent(lobjDocument, lobjZipFile)

          Return lobjDocument

        Else
          Return Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ' <Modified by: Ernie at 9/20/2011-8:23:18 AM on machine: ERNIE-M4400>
    ' We are not able to support related documents with Silverlight at this time
#If Not SILVERLIGHT = 1 Then

#End If
    ' </Modified by: Ernie at 9/20/2011-8:23:18 AM on machine: ERNIE-M4400>

    Private Sub ExtractContent(ByVal lpDocument As Document, ByVal lpZipFile As ZipFile)
      Try

        If lpZipFile.Entries.Count > 0 Then

          'Dim lobjCurrentZipEntry As ZipEntry
          Dim lobjContentStream As Stream
          Dim lstrRelativePath As String = String.Empty

          For Each lobjVersion As Version In lpDocument.Versions
            For Each lobjContent As Content In lobjVersion.Contents
              If lobjContent.FileSize.Megabytes < ConfigurationManager.AppSettings("MaximumInMemoryDocumentMegabytes") Then
                lobjContentStream = New MemoryStream
              Else
                Dim lstrTempFilePath As String = Path.GetTempFileName()
                lobjContentStream = New FileStream(lstrTempFilePath, FileMode.Open)
                mobjTempFilePaths.Add(lstrTempFilePath)
                mobjTempFileStreams.Add(lobjContentStream)
              End If

              'lobjZipFile.Extract(lobjContent.RelativePath.Replace("\", "/"), lobjContentStream)
              If String.IsNullOrEmpty(lobjContent.RelativePath) Then
                lstrRelativePath = String.Format("{0}/{1}", lobjVersion.ID, lobjContent.FileName)
              Else
                lstrRelativePath = lobjContent.RelativePath.Replace("\", "/")
              End If
              If lpZipFile.EntryFileNames.Contains(lstrRelativePath) Then
                lpZipFile.Item(lstrRelativePath).Extract(lobjContentStream)
                With lobjContent
                  '.StorageType = Content.StorageTypeEnum.EncodedUnCompressed
                  .Data.SetStream(lobjContentStream)
                End With
              Else
                ApplicationLogging.WriteLogEntry(String.Format("Missing entry '{0}' in archive '{1}'.", lstrRelativePath,
                                                               lpZipFile.Info))
              End If

              ' Check for annotations
              lstrRelativePath = $"{lobjVersion.ID}/annotations"
              If ZipContainsAnnotations(lpZipFile, lstrRelativePath) Then
                Dim lintAnnotationCounter As Integer = 0
                Dim lobjAnnotationStream As Stream
                For Each lstrEntryFileName As String In lpZipFile.EntryFileNames
                  If lstrEntryFileName.Contains(lstrRelativePath) Then
                    lobjAnnotationStream = New MemoryStream
                    lpZipFile.Item(lstrEntryFileName).Extract(lobjAnnotationStream)
                    lobjContent.Annotations.Item(lintAnnotationCounter).AnnotatedContent = New Content(New NamedStream(lobjAnnotationStream, Path.GetFileName(lstrEntryFileName)))
                    lintAnnotationCounter += 1
                  End If
                Next
              End If

            Next
          Next

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function ZipContainsAnnotations(ByRef lpZipFile As ZipFile, Optional lpRelativePath As String = "0/annotations")
      Try
        For Each lstrEntryFileName As String In lpZipFile.EntryFileNames
          If lstrEntryFileName.Contains(lpRelativePath) Then
            Return True
          End If
        Next
        Return False
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Saves the current instance to an archive
    ''' </summary>
    ''' <param name="archivePath">The fully qualified path to save the archive to.</param>
    ''' <remarks></remarks>
    Public Sub Archive(ByVal archivePath As String) _
      Implements IArchivable(Of Document).Archive
      Try
        Archive(archivePath, True, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Saves the current instance to an archive
    ''' </summary>
    ''' <param name="archivePath">The fully qualified path to save the archive to.</param>
    ''' <remarks></remarks>
    Public Sub Archive(ByVal archivePath As String, ByVal removeContainedFiles As Boolean) _
      Implements IArchivable(Of Document).Archive
      Try
        Archive(archivePath, removeContainedFiles, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Saves the current instance to an archive
    ''' </summary>
    ''' <param name="archivePath">The fully qualified path to save the archive to.</param>
    ''' <remarks></remarks>
    Public Sub Archive(ByVal archivePath As String, ByVal removeContainedFiles As Boolean, ByVal password As String) _
      Implements IArchivable(Of Document).Archive
      Try
        ' BMK Archive
        ' Verify the archive path
        Helper.VerifyFilePath(archivePath, False)

        If SerializationPath Is Nothing OrElse SerializationPath.Length = 0 Then
          SerializationPath = GetSerializationPathFromArchivePath(archivePath)
          mstrCurrentPath = SerializationPath
        End If

        Dim lobjArchiveStream As Stream = Nothing

        If IO.Path.GetExtension(archivePath).Replace(".", String.Empty) = JSON_CONTENT_PACKAGE_FILE_EXTENSION Then
          lobjArchiveStream = ToArchiveStream(password, True)
        Else
          lobjArchiveStream = ToArchiveStream(password, False)
        End If

        Dim lobjArchiveFileStream As New FileStream(archivePath, FileMode.Create, FileAccess.Write)

        lobjArchiveStream.CopyTo(lobjArchiveFileStream)

        'If TypeOf lobjArchiveStream Is MemoryStream Then
        '  DirectCast(lobjArchiveStream, MemoryStream).WriteTo(lobjArchiveFileStream)
        'ElseIf TypeOf lobjArchiveStream Is FileStream Then
        '  lobjArchiveStream.CopyTo(lobjArchiveFileStream)
        'End If

        lobjArchiveFileStream.Close()

        lobjArchiveStream.Dispose()
        lobjArchiveFileStream.Dispose()

        If removeContainedFiles = True Then
          ' We need to clean up the constituent files
          'DeleteContentPaths()
          DeleteAllRelatedFiles(Path.GetDirectoryName(archivePath))
        End If

        'If mobjTempContentFileStream IsNot Nothing Then
        '  mobjTempContentFileStream.Dispose()
        'End If

        'If IO.File.Exists(mstrTempContentFilePath) Then
        '  IO.File.Delete(mstrTempContentFilePath)
        'End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function GetSerializationPathFromArchivePath(ByVal lpArchivePath As String) As String
      Try

        Dim lstrArchiveFileName As String
        Dim lstrSerializationFolder As String
        Dim lstrSerializationPath As String

        lstrArchiveFileName = Path.GetFileNameWithoutExtension(lpArchivePath)
        lstrSerializationFolder = Path.GetDirectoryName(lpArchivePath)
        lstrSerializationPath = String.Format("{0}\{1}.{2}", lstrSerializationFolder, lstrArchiveFileName,
                                              CONTENT_DEFINITION_FILE_EXTENSION)

        Return lstrSerializationPath

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Saves the current instance to a ZipFile stream.
    ''' </summary>
    ''' <returns>An IO.Stream of a DotNetZip ZipFile stream.</returns>
    ''' <remarks></remarks>
    Public Function ToArchiveStream() As IO.Stream _
      Implements IArchivable(Of Core.Document).ToArchiveStream
      Try
        Return ToArchiveStream(String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToJsonArchiveStream() As IO.Stream
      Try
        Return ToArchiveStream(String.Empty, True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToArchiveStream(ByVal password As String) As IO.Stream _
      Implements IArchivable(Of Core.Document).ToArchiveStream
      Try
        Return ToArchiveStream(password, False)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Saves the current instance to a ZipFile stream.
    ''' </summary>
    ''' <param name="password">A password to set on the zipped document.  
    ''' If no password is wanted then simply pass null or String.Empty.</param>
    ''' <returns>An IO.Stream of a DotNetZip ZipFile stream.</returns>
    ''' <remarks></remarks>
    Public Function ToArchiveStream(ByVal password As String, metadataAsJson As Boolean) As IO.Stream
      Try

        If Me.TotalContentSize.Megabytes < ConfigurationManager.AppSettings("MaximumInMemoryDocumentMegabytes") Then
          Return ToArchiveStream(New MemoryStream, password, metadataAsJson)
        Else
          Dim lstrTempFilePath As String = Path.GetTempFileName()
          Dim lobjTempContentFileStream As New FileStream(lstrTempFilePath, FileMode.Open)
          mobjTempFilePaths.Add(lstrTempFilePath)
          mobjTempFileStreams.Add(lobjTempContentFileStream)
          Return ToArchiveStream(lobjTempContentFileStream, password, metadataAsJson)

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToArchiveStream(ByVal lpOutputStream As Stream, ByVal password As String, metadataAsJson As Boolean) _
      As IO.Stream
      Try

        Dim lobjAnnotationEntry As ZipFileStream
        Dim lobjContentEntry As ZipFileStream
        Dim lobjAnnotationStream As Stream
        Dim lobjContentStream As Stream
        Dim lobjContentStreams As New List(Of ZipFileStream)
        Dim lobjAnnotationStreams As New List(Of ZipFileStream)
        Dim lobjZipFile As New ZipFile
        Dim lobjDocumentStream As Stream = Nothing

        ' We need to read the content files before we read the document object itself.  
        ' Otherwise the content file streams will be wiped out during the reserialization 
        ' that occurs during UpdateHeader.  UpdateHeader is triggered when serializing the 
        ' document object to a stream.

        ' However, we need to add the document object to the zip file first for our defined 
        ' zip file structure.  In order to accomplish this balancing act we will cache the 
        ' related documents and content streams in a list first.  After we add the document 
        ' stream to the zip file we will read back the content streams from the list and add 
        ' them to the zip file.  
        ' In this way we can order the zip entries the way we need to while reading the content 
        ' and document information in the sequence needed.

        ' If at some point in the future we can prevent UpdateHeader from clearing out our 
        ' content streams then we will no longer need to perform these acrobatics to properly 
        ' create the archive stream.
        '
        ' Ernie Bahr
        ' 8/11/2010


        With lobjZipFile

          .ParallelDeflateThreshold = -1

          ' Get the relationships
          lobjContentStreams.AddRange(GetRelationshipStreams(password))

          ' Get the versions
          For Each lobjVersion As Version In Me.Versions
            For Each lobjContent As Content In lobjVersion.Contents
              If lobjContent.Data.Length > 0 Then
                lobjContentStream = lobjContent.Data.ToStream()
                If lobjContent.FileSize.Megabytes < ConfigurationManager.AppSettings("MaximumInMemoryDocumentMegabytes") Then
                  lobjContentEntry = New ZipFileStream(lobjContent.FileName, lobjVersion.ID.ToString,
                                                       Helper.CopyStream(lobjContentStream))
                Else
                  ApplicationLogging.WriteLogEntry(String.Format(
                    "Unable to archive content file '{0}' of size {1} using memory stream, trying file stream instead.",
                    lobjContent.FileName, lobjContent.FileSize.ToString()),
                                                   Reflection.MethodBase.GetCurrentMethod, TraceEventType.Information,
                                                   63456)
                  ' Try to use a file stream instead
                  Dim lstrTempFilePath As String = String.Empty
                  Dim lobjTempContentFileStream As FileStream
                  lobjTempContentFileStream = Helper.CopyStreamToTempFileStream(lobjContentStream, lstrTempFilePath)
                  lobjContentEntry = New ZipFileStream(lobjContent.FileName, lobjVersion.ID.ToString,
                                                       lobjTempContentFileStream)
                  mobjTempFilePaths.Add(lstrTempFilePath)
                  mobjTempFileStreams.Add(lobjTempContentFileStream)
                End If

                If lobjContent.Annotations IsNot Nothing AndAlso lobjContent.Annotations.Count > 0 Then
                  For Each lobjAnnotation As Annotation In lobjContent.Annotations
                    If lobjAnnotation.AnnotatedContent IsNot Nothing AndAlso lobjAnnotation.AnnotatedContent.Data.Length > 0 Then
                      lobjAnnotationStream = lobjAnnotation.AnnotatedContent.Data.ToStream()
                      lobjAnnotationEntry = New ZipFileStream(lobjAnnotation.AnnotatedContent.FileName, $"{lobjVersion.ID.ToString}\annotations\{lobjAnnotation.ID}",
                                                       Helper.CopyStream(lobjAnnotationStream))
                    End If
                    If lobjAnnotationEntry IsNot Nothing Then
                      lobjAnnotationStreams.Add(lobjAnnotationEntry)
                    End If
                  Next
                End If

                If lobjContentEntry IsNot Nothing Then
                  lobjContentStreams.Add(lobjContentEntry)
                End If
              End If
            Next
          Next

          SetDocumentReferences()

          UpdateHeader()
          SetContentPathsFromSerializationPath()

          Dim lstrCleanFileName As String

          If metadataAsJson Then
            lobjDocumentStream = Me.ToJsonStream()

            lstrCleanFileName = Helper.CleanFile(String.Format("{0}.{1}", Me.ID, JSON_CONTENT_DEFINITION_FILE_EXTENSION),
                                                 "^")
            ' lstrCleanFileName = String.Format("{0}.{1}", Me.ID, JSON_CONTENT_DEFINITION_FILE_EXTENSION)
          Else
            lobjDocumentStream = Serializer.Serialize.ToStream(Me)
            lstrCleanFileName = Helper.CleanFile(String.Format("{0}.{1}", Me.ID, CONTENT_DEFINITION_FILE_EXTENSION), "^")
          End If

          '#If CTSDOTNET = 1 Then
          '          lstrCleanFileName = Helper.CleanFile(String.Format("{0}.{1}", Me.ID, CONTENT_DEFINITION_FILE_EXTENSION), "^")
          '#Else
          '          lstrCleanFileName = String.Format("{0}.{1}", Me.ID, CONTENT_DEFINITION_FILE_EXTENSION)
          '#End If

          'If lobjDocumentStream Is Nothing Then
          '  'Throw New Exceptions.do
          'End If

          ' If supplied, set the password
          If password IsNot Nothing AndAlso password.Length > 0 Then
            lobjZipFile.Password = password
          End If

          '.AddFileStream(String.Format("{0}.{1}", Me.ID, CONTENT_DEFINITION_FILE_EXTENSION), String.Empty, lobjDocumentStream)


          .AddEntry(lstrCleanFileName, lobjDocumentStream)

          For Each lobjZipContentStream As ZipFileStream In lobjContentStreams
            '.AddFileStream(lobjContentStream.FileName, lobjContentStream.DirectoryInArchive, lobjContentStream.Stream)
            .AddEntry(String.Format("{0}\{1}", lobjZipContentStream.DirectoryInArchive,
                                    lobjZipContentStream.FileName), lobjZipContentStream.Stream)
            lobjZipContentStream = Nothing
          Next

          For Each lobjZipAnnotationStream As ZipFileStream In lobjAnnotationStreams
            '.AddFileStream(lobjContentStream.FileName, lobjContentStream.DirectoryInArchive, lobjContentStream.Stream)
            .AddEntry(String.Format("{0}\{1}", lobjZipAnnotationStream.DirectoryInArchive,
                                    lobjZipAnnotationStream.FileName), lobjZipAnnotationStream.Stream)
            lobjZipAnnotationStream = Nothing
          Next

        End With

        Try
          lobjZipFile.Save(lpOutputStream)
        Catch MemEx As OutOfMemoryException
          ApplicationLogging.LogException(MemEx, Reflection.MethodBase.GetCurrentMethod)
          ApplicationLogging.WriteLogEntry("Unable to save zip file to stream, the file is too large.")
          '  Re-throw the exception to the caller
          Throw
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try


        If lpOutputStream.CanSeek Then
          lpOutputStream.Position = 0
        End If

        lobjZipFile.Dispose()

        lobjZipFile = Nothing

        Return lpOutputStream

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function


    Public Function ToArchiveBytes() As Byte()
      Try
        Dim lobjArchiveStream As Stream = ToArchiveStream()
        Dim lobjReturnBytes As Byte() = Helper.ReadStreamToByteArray(lobjArchiveStream)
        lobjArchiveStream.Dispose()
        Return lobjReturnBytes
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Function GetRelationshipStreams(ByVal password As String) As List(Of ZipFileStream)
      Try
        Dim lobjContentStreams As New List(Of ZipFileStream)

        If Me.Relationships IsNot Nothing AndAlso Me.Relationships.Count > 0 Then
          Dim lobjRelatedDocument As Document = Nothing
          Dim lstrRelatedDocumentArchiveFileName As String = String.Empty
          Dim lblnEntryExists As Boolean
          Dim lobjRelatedDocumentStream As Stream

          ' Get the relationships
          For Each lobjRelationship As Relationship In Me.Relationships
            If lobjRelationship.RelatedDocument IsNot Nothing Then
              lobjRelatedDocument = lobjRelationship.RelatedDocument
              If lobjRelatedDocument.SerializationPath Is Nothing OrElse
                 lobjRelatedDocument.SerializationPath = String.Empty Then
                lobjRelatedDocument.SerializationPath = String.Format("{0}\{1}.cdf",
                                                                      Me.SerializationPath.Replace(".cdf", String.Empty),
                                                                      lobjRelatedDocument.ID)
              End If
              lstrRelatedDocumentArchiveFileName = String.Format("{0}.{1}", lobjRelatedDocument.ID,
                                                                 CONTENT_PACKAGE_FILE_EXTENSION)
              lblnEntryExists = False

              For Each lobjRelationshipStream As ZipFileStream In lobjContentStreams
                If lobjRelationshipStream.FileName = lstrRelatedDocumentArchiveFileName Then
                  lblnEntryExists = True
                End If
              Next

              If lblnEntryExists = False Then
                lobjRelatedDocumentStream = lobjRelatedDocument.ToArchiveStream(password)
                lobjContentStreams.Add(New ZipFileStream(lstrRelatedDocumentArchiveFileName, "Relationships",
                                                         lobjRelatedDocumentStream))
              End If
            End If
          Next
        End If

        Return lobjContentStreams

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class
End Namespace