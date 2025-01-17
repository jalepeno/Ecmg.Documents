'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

''' <summary>
''' Contains the defined exceptions of the Cts framework
''' </summary>
Namespace Exceptions

  ''' <summary>
  ''' Exception thrown as a result of a document 
  ''' failing validation before importing
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentValidationException
    Inherits DocumentException

#Region "Class Variables"

    Private m_documentId As String
    Private m_versionId As String
    Private m_Document As Document

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
    ''' Creates a new DocumentException with the specified document id.
    ''' </summary>
    ''' <param name="id">The Document Id</param>
    ''' <remarks>Initializes the VersionId to zero and the document object to a null object reference</remarks>
    Public Sub New(ByVal id As String, ByVal message As String, ByVal innerException As System.Exception)
      Me.New(id, "0", message)
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document.
    ''' </summary>
    ''' <param name="document">The Document object</param>
    ''' <remarks>Initializes the VersionId to zero</remarks>
    Public Sub New(ByVal document As Document, ByVal message As String)
      Me.New(document, "0", message)
    End Sub

    ''' <summary>
    ''' Creates a new DocumentException with the specified document id and version id.
    ''' </summary>
    ''' <param name="id">The Document Id</param>
    ''' <param name="versionId">The Version Id</param>
    ''' <remarks>Initializes the document object to a null object reference</remarks>
    Public Sub New(ByVal id As String, ByVal versionId As String, ByVal message As String)
      MyBase.New(message, CType(Nothing, Exception))
      Try
        m_documentId = id
        m_versionId = versionId
        HResult = 16468
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
    Public Sub New(ByVal document As Document, ByVal versionId As String, ByVal message As String)
      Me.New(document, versionId, message, Nothing)
    End Sub

    Public Sub New(ByVal document As Document, ByVal versionId As String, ByVal message As String, ByVal innerException As System.Exception)

      MyBase.New(message, innerException)
      Try
        m_Document = document
        m_documentId = document.ID
        m_versionId = versionId
        HResult = 16468
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try
    End Sub
#End Region

  End Class
End Namespace