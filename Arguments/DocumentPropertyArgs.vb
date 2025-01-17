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

Namespace Arguments

  ''' <summary>
  ''' Contains all the necessary parameters for the UpdateDocument method
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentPropertyArgs
    Inherits ObjectPropertyArgs

#Region "Class Variables"

    Private mstrVersionID As String = String.Empty
    Private menuVersionScope As VersionScopeEnum
    Private mobjProperties As ECMProperties = New ECMProperties



#End Region
#Region "Public Properties"
    ''' <summary>
    ''' Identifies the document to be updated
    ''' </summary>
    Public Property DocumentID() As String
      Get
        Try
          Return ObjectID
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          ObjectID = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Identifies which version should be updated
    ''' </summary>
    Public Property VersionID() As String
      Get
        Try
          Return mstrVersionID
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          mstrVersionID = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Alternative way to determine which version should be updated.
    ''' </summary>
    Public Property VersionScope As VersionScopeEnum
      Get
        Try
          Return menuVersionScope
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As VersionScopeEnum)
        Try
          menuVersionScope = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentUpdateArgs
    ''' </summary>
    ''' <param name="lpDocumentID">The document identifier</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpDocumentID As String)
      Me.New(lpDocumentID, "", New ECMProperties)
    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentUpdateArgs
    ''' </summary>
    ''' <param name="lpDocumentID">The document identifier</param>
    ''' <param name="lpVersionID">The version identifier</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpDocumentID As String, ByVal lpVersionID As String)
      Me.New(lpDocumentID, lpVersionID, New ECMProperties)
    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentUpdateArgs
    ''' </summary>
    ''' <param name="lpDocumentID">The document identifier</param>
    ''' <param name="lpVersionID">The version identifier</param>
    ''' <param name="lpProperties">A Core.Properties object containing the properties to update</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpDocumentID As String, ByVal lpVersionID As String, ByVal lpProperties As ECMProperties)

      Try
        DocumentID = lpDocumentID
        VersionID = lpVersionID
        Properties = lpProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentUpdateArgs
    ''' </summary>
    ''' <param name="lpDocumentID">The document identifier</param>
    ''' <param name="lpVersionScope">The version scope</param>
    ''' <param name="lpProperties">A Core.Properties object containing the properties to update</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpDocumentID As String, ByVal lpVersionScope As VersionScopeEnum, ByVal lpProperties As ECMProperties)

      Try

        DocumentID = lpDocumentID
        VersionScope = lpVersionScope
        Properties = lpProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

#End Region

  End Class

End Namespace