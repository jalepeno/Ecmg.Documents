'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters necessary for the ExportDocument method.
  ''' </summary>
  ''' <remarks>Used as a sole parameter for the ExportDocument method.</remarks>
  Public Class ExportDocumentEventArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mstrId As String
    Private mobjDocument As Core.Document
    Private mobjFolderContent As FolderContent
    Private mobjStorageType As Content.StorageTypeEnum = Content.StorageTypeEnum.Reference
    Private mblnGetContent As Boolean = True
    Private mblnGetContentFileNames As Boolean = True
    Private mblnGetPermissions As Boolean = True
    Private mblnGenerateCDF As Boolean = False
    Private mstrTransformationPath As String = String.Empty 'Path to a transformation file
    Private mobjTransformation As Transformations.Transformation = Nothing 'Transformation object
    Private mblnGetRecord As Boolean = False
    Private mobjRecordContentSource As Providers.ContentSource
    Private mblnGetAnnotations As Boolean = True
    Private mblnArchive As Boolean = False
    Private mobjVersionScope As New ExportVersionScope(VersionScopeEnum.AllVersions)
    Private mblnGetRelatedDocuments As Boolean = True
    Private mobjPropertyFilter As List(Of String) = Nothing

#End Region

#Region "Public Properties"

    Public ReadOnly Property Id() As String
      Get
        Return mstrId
      End Get
    End Property

    Public Property Document() As Core.Document
      Get
        Return mobjDocument
      End Get
      Set(ByVal value As Core.Document)
        Try
          mobjDocument = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property VersionScope As ExportVersionScope
      Get
        Return mobjVersionScope
      End Get
      Set(value As ExportVersionScope)
        Try
          mobjVersionScope = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property FolderContent() As FolderContent
      Get
        Return mobjFolderContent
      End Get
      Set(ByVal value As FolderContent)
        Try
          mobjFolderContent = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property StorageType() As Content.StorageTypeEnum
      Get
        Return mobjStorageType
      End Get
      Set(ByVal value As Content.StorageTypeEnum)
        mobjStorageType = value
      End Set
    End Property

    ''' <summary>
    ''' Determines whether or not the export operation 
    ''' should get or exclude the document content
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Value of true indicates the export 
    ''' should get the document content, 
    ''' value of False indicates it should export 
    ''' the document without the content.</remarks>
    Public Property GetContent() As Boolean
      Get
        Return mblnGetContent
      End Get
      Set(ByVal value As Boolean)
        mblnGetContent = value
      End Set
    End Property

    ''' <summary>
    ''' Determines whether or not the export operation 
    ''' should get or exclude the document content file name(s)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Value of true indicates the export 
    ''' should get the document content file name(s), 
    ''' value of False indicates it should export 
    ''' the document without the file name(s).</remarks>
    Public Property GetContentFileNames() As Boolean
      Get
        Return mblnGetContentFileNames
      End Get
      Set(ByVal value As Boolean)
        mblnGetContentFileNames = value
      End Set
    End Property

    ''' <summary>Determines whether or not the export operation should get or exclude the document permissions</summary>
    ''' <remarks>This property can only be used on exports with providers that implement ISecurityClassification.</remarks>
    ''' <seealso cref="Providers.ISecurityClassification">ISecurityClassification</seealso>
    Public Property GetPermissions As Boolean
      Get
        Return mblnGetPermissions
      End Get
      Set(value As Boolean)
        mblnGetPermissions = value
      End Set
    End Property

    ''' <summary>
    ''' GenerateCDF
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GenerateCDF() As Boolean
      Get
        Return mblnGenerateCDF
      End Get
      Set(ByVal value As Boolean)
        mblnGenerateCDF = value
      End Set
    End Property

    Public Property TransformationPath() As String
      Get
        Return mstrTransformationPath
      End Get
      Set(ByVal value As String)
        Try
          mobjTransformation = New Transformations.Transformation(value)
        Catch ex As Exception
          Throw New Exception(ex.Message)
        End Try
        mstrTransformationPath = value
      End Set
    End Property

    Public Property Transformation() As Transformations.Transformation
      Get
        Return mobjTransformation
      End Get
      Set(ByVal value As Transformations.Transformation)
        Try
          mobjTransformation = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Determines whether or not the export operation 
    ''' should get or exclude related record if it exists
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GetRecord() As Boolean
      Get
        Return mblnGetRecord
      End Get
      Set(ByVal value As Boolean)
        mblnGetRecord = value
      End Set
    End Property

    ''' <summary>
    ''' Determines whether or not the export operation 
    ''' should get or exclude related documents
    ''' </summary>
    ''' <returns></returns>
    Public Property GetRelatedDocuments As Boolean
      Get
        Return mblnGetRelatedDocuments
      End Get
      Set(ByVal value As Boolean)
        mblnGetRelatedDocuments = value
      End Set
    End Property

    ''' <summary>
    ''' If GetRecord is true, then export will use this ContentSource
    ''' to go find the related record (if exists) and return it in a 
    ''' Relationship object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RecordContentSource() As Providers.ContentSource
      Get
        Return mobjRecordContentSource
      End Get
      Set(ByVal value As Providers.ContentSource)
        Try
          mobjRecordContentSource = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property GetAnnotations As Boolean
      Get
        Return mblnGetAnnotations
      End Get
      Set(ByVal value As Boolean)
        mblnGetAnnotations = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not the export 
    ''' should write the document out as an Archive.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>
    ''' Writing out the document as an archive will 
    ''' write it to a single zip style archive file 
    ''' containing both the original cdf and all of 
    ''' the content files.  This is only respected 
    ''' when used in conjunction with a GetContent 
    ''' value of True.
    ''' </remarks>
    Public Property Archive As Boolean
      Get
        Return mblnArchive
      End Get
      Set(ByVal value As Boolean)
        mblnArchive = value
      End Set
    End Property


    Public Property PropertyFilter As List(Of String)
      Get
        Return mobjPropertyFilter
      End Get
      Set(value As List(Of String))
        mobjPropertyFilter = value
      End Set
    End Property
#End Region

#Region "Constructors"

    Public Sub New(ByVal lpId As String)
      Me.New(lpId, Nothing, Nothing)
    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpDocument As Core.Document)
      Me.New(lpId, lpDocument, Nothing)
    End Sub

    Public Sub New(ByVal lpId As String,
                   ByVal lpWorker As BackgroundWorker)

      Me.New(lpId, Nothing, lpWorker)

    End Sub

    Public Sub New(ByVal lpId As String,
                   ByRef lpDocument As Core.Document,
                   ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpWorker)

      Try
        mstrId = lpId
        mobjDocument = lpDocument
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Sub New(ByVal lpFolderContent As FolderContent)
      Me.New(lpFolderContent.ID, Nothing, Nothing)
      Try
        mobjFolderContent = lpFolderContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFolderContent As FolderContent, ByVal lpDocument As Core.Document)
      Me.New(lpFolderContent.ID, lpDocument, Nothing)
      Try
        mobjFolderContent = lpFolderContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFolderContent As FolderContent, ByVal lpWorker As BackgroundWorker)
      Me.New(lpFolderContent.ID, Nothing, lpWorker)
      Try
        mobjFolderContent = lpFolderContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFolderContent As FolderContent, ByVal lpWorker As BackgroundWorker, ByVal lpTransformation As Transformations.Transformation)
      Me.New(lpFolderContent.ID, Nothing, lpWorker)
      Try
        mobjFolderContent = lpFolderContent
        mobjTransformation = lpTransformation
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFolderContent As FolderContent, ByVal lpDocument As Core.Document, ByVal lpWorker As BackgroundWorker, ByVal lpTransformation As Transformations.Transformation)
      Me.New(lpFolderContent.ID, lpDocument, lpWorker)
      Try
        mobjFolderContent = lpFolderContent
        mobjTransformation = lpTransformation
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpFolderContent As FolderContent, ByVal lpDocument As Core.Document, ByVal lpWorker As BackgroundWorker)
      Me.New(lpFolderContent.ID, lpDocument, lpWorker)
      Try
        mobjFolderContent = lpFolderContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpId As String,
                   ByVal lpGetContent As Boolean)
      Me.New(lpId, Nothing, Nothing)
      Try
        StorageType = Content.StorageTypeEnum.Reference
        GetContent = lpGetContent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
