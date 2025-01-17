'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  FolderErrorEventArgs.vb
'   Description :  [type_description_here]
'   Created     :  3/6/2015 9:48:37 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Utilities


#End Region

Namespace Arguments

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class FolderErrorEventArgs
    Inherits ErrorEventArgs

#Region "Class Variables"

    Private mstrFolderId As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property FolderId() As String
      Get
        Return mstrFolderId
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

    Public Sub New(ByVal lpExportArgs As ExportFolderEventArgs,
                   ByVal lpException As Exception)
      MyBase.New(String.Format("ID:{0}:{1}", lpExportArgs.PrimaryIdentifier, lpException.Message), lpException)
      Try
        mstrFolderId = lpExportArgs.PrimaryIdentifier
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportFolderEventArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(String.Format("ID:{0}:{1}", lpExportArgs.PrimaryIdentifier, lpException.Message), lpException, lpWorker)
      Try
        mstrFolderId = lpExportArgs.PrimaryIdentifier
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpExportArgs As ImportFolderArgs,
                   ByVal lpException As Exception)
      MyBase.New(String.Format("Name:{0}:{1}", lpExportArgs.Folder.Name, lpException.Message), lpException)
      Try
        If lpExportArgs.Folder IsNot Nothing Then
          If Not String.IsNullOrEmpty(lpExportArgs.Folder.Id) Then
            mstrFolderId = lpExportArgs.Folder.Id
          ElseIf Not String.IsNullOrEmpty(lpExportArgs.Folder.Name) Then
            mstrFolderId = lpExportArgs.Folder.Name
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpExportArgs As ImportFolderArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(String.Format("Name:{0}:{1}", lpExportArgs.Folder.Name, lpException.Message), lpException, lpWorker)
      Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        If lpExportArgs.Folder IsNot Nothing Then
          If Not String.IsNullOrEmpty(lpExportArgs.Folder.Id) Then
            mstrFolderId = lpExportArgs.Folder.Id
          ElseIf Not String.IsNullOrEmpty(lpExportArgs.Folder.Name) Then
            mstrFolderId = lpExportArgs.Folder.Name
          End If
        End If
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpSourceFolderId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpMessage, lpException)
      mstrFolderId = lpSourceFolderId
    End Sub

    Public Sub New(ByVal lpSourceFolderId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpMessage, lpException, lpWorker)
      mstrFolderId = lpSourceFolderId
    End Sub

#End Region


#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        If FolderId.Length > 0 Then
          Return String.Format("FolderId={0}; {1}", FolderId, MyBase.DebuggerIdentifier)
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