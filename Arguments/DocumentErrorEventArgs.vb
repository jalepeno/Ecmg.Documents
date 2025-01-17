'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Utilities


#End Region

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the ExportError Event
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class DocumentErrorEventArgs
    Inherits ErrorEventArgs

#Region "Class Variables"

    Private mstrDocumentId As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property DocumentId() As String
      Get
        Return mstrDocumentId
      End Get
    End Property

#End Region

#Region "Constructors"

    ' RaiseEvent DocumentExportError(Me, New ExportErrorEventArgs(Args.Id, String.Format(Args.Id, "ID: {0} :{1}", Args.Id, ex.Message), ex))
    ' Where Args is ExportDocumentEventArgs

    Public Sub New(ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpMessage, lpException)
    End Sub

    Public Sub New(ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpMessage, lpException, lpWorker)
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportDocumentEventArgs,
                   ByVal lpException As Exception)
      MyBase.New(String.Format("ID:{0}:{1}", lpExportArgs.Id, lpException.Message), lpException)
      Try
        mstrDocumentId = lpExportArgs.Id
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportDocumentEventArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(String.Format("ID:{0}:{1}", lpExportArgs.Id, lpException.Message), lpException, lpWorker)
      Try
        mstrDocumentId = lpExportArgs.Id
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpExportArgs As ImportDocumentArgs,
                   ByVal lpException As Exception)
      MyBase.New(String.Format("Name:{0}:{1}", lpExportArgs.Document.Name, lpException.Message), lpException)
      Try
        If lpExportArgs.Document IsNot Nothing Then
          If Not String.IsNullOrEmpty(lpExportArgs.Document.ID) Then
            mstrDocumentId = lpExportArgs.Document.ID
          ElseIf Not String.IsNullOrEmpty(lpExportArgs.Document.Name) Then
            mstrDocumentId = lpExportArgs.Document.Name
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpExportArgs As ImportDocumentArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(String.Format("Name:{0}:{1}", lpExportArgs.Document.Name, lpException.Message), lpException, lpWorker)
      If lpExportArgs.Document IsNot Nothing Then
        If Not String.IsNullOrEmpty(lpExportArgs.Document.ID) Then
          mstrDocumentId = lpExportArgs.Document.ID
        ElseIf Not String.IsNullOrEmpty(lpExportArgs.Document.Name) Then
          mstrDocumentId = lpExportArgs.Document.Name
        End If
      End If
    End Sub

    Public Sub New(ByVal lpSourceDocumentId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpMessage, lpException)
      mstrDocumentId = lpSourceDocumentId
    End Sub

    Public Sub New(ByVal lpSourceDocumentId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpMessage, lpException, lpWorker)
      mstrDocumentId = lpSourceDocumentId
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        If DocumentId.Length > 0 Then
          Return String.Format("DocumentId={0}; {1}", DocumentId, MyBase.DebuggerIdentifier)
        Else
          Return MyBase.DebuggerIdentifier
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace