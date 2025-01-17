'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ObjectExportErrorEventArgs.vb
'   Description :  [type_description_here]
'   Created     :  9/1/2015 3:29:01 AM
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

  Public Class ObjectExportErrorEventArgs
    Inherits ObjectErrorEventArgs

#Region "Public Properties"

    Public ReadOnly Property SourceObjectId() As String
      Get
        Return MyBase.ObjectId
      End Get
    End Property

#End Region

#Region "Constructors"

    ' RaiseEvent FolderExportError(Me, New ExportErrorEventArgs(Args.Id, String.Format(Args.Id, "ID: {0} :{1}", Args.Id, ex.Message), ex))
    ' Where Args is ExportFolderEventArgs

    Public Sub New(ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpMessage, lpException)
    End Sub

    Public Sub New(ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpMessage, lpException, lpWorker)
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportObjectEventArgs,
                   ByVal lpException As Exception)
      MyBase.New(lpExportArgs, lpException)
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportObjectEventArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpExportArgs, lpException, lpWorker)
    End Sub

    Public Sub New(ByVal lpSourceObjectId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpSourceObjectId, lpMessage, lpException)
    End Sub

    Public Sub New(ByVal lpSourceObjectId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpSourceObjectId, lpMessage, lpException, lpWorker)
    End Sub

#End Region

#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        If SourceObjectId.Length > 0 Then
          Return String.Format("SourceObjectId={0}; {1}", SourceObjectId, MyBase.DebuggerIdentifier)
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