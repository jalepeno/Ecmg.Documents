'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

''' <summary>
''' Contains the defined exceptions of the Cts framework
''' </summary>
Namespace Exceptions

  ''' <summary>
  ''' Exception thrown by a document or from methods associated with documents
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentException
    Inherits CtsException

#Region "Class Variables"

    Private m_documentId As String
    Private m_versionId As String
    Private m_Document As Core.Document

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The Document Identifier
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property DocumentId() As String
      Get
        Return m_documentId
      End Get
    End Property

    ''' <summary>
    ''' The Version Identifier
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property VersionId() As String
      Get
        Return m_versionId
      End Get
    End Property

    ''' <summary>
    ''' Reference to the Document object with which the exception is associated
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Document() As Core.Document
      Get
        Return m_Document
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document id.
    ''' </summary>
    ''' <param name="id">The Document Id</param>
    ''' <remarks>Initializes the VersionId to zero and the document object to a null object reference</remarks>
    Public Sub New(ByVal id As String, ByVal message As String)
      Me.New(id, "0", message)
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified message and inner exception.
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="innerException"></param>
    ''' <remarks></remarks>
    Protected Sub New(ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document id.
    ''' </summary>
    ''' <param name="id">The Document Id</param>
    ''' <remarks>Initializes the VersionId to zero and the document object to a null object reference</remarks>
    Public Sub New(ByVal id As String, ByVal message As String, ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
      m_documentId = id
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document.
    ''' </summary>
    ''' <param name="document">The Document object</param>
    ''' <remarks>Initializes the VersionId to zero</remarks>
    Public Sub New(ByVal document As Core.Document, ByVal message As String)
      Me.New(document, "0", message)
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document id and version id.
    ''' </summary>
    ''' <param name="id">The Document Id</param>
    ''' <param name="versionId">The Version Id</param>
    ''' <remarks>Initializes the document object to a null object reference</remarks>
    Public Sub New(ByVal id As String, ByVal versionId As String, ByVal message As String)
      MyBase.New(message, Nothing)
      Try
        m_documentId = id
        m_versionId = versionId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document.
    ''' </summary>
    ''' <param name="document">The Document object</param>
    ''' <param name="versionId">The Version Id</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal document As Core.Document, ByVal versionId As String, ByVal message As String)
      Me.New(document, versionId, message, Nothing)
    End Sub

    Public Sub New(ByVal document As Core.Document, ByVal versionId As String, ByVal message As String, ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
      Try
        m_Document = document
        m_documentId = document.ID
        m_versionId = versionId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub
#End Region

  End Class

End Namespace