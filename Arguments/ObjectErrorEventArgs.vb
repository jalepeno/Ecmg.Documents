'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ObjectErrorEventArgs.vb
'   Description :  [type_description_here]
'   Created     :  9/1/2015 3:27:37 AM
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
  Public Class ObjectErrorEventArgs
    Inherits ErrorEventArgs

#Region "Class Variables"

    Private mstrObjectId As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ObjectId() As String
      Get
        Return mstrObjectId
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

    Public Sub New(ByVal lpExportArgs As ExportObjectEventArgs,
                   ByVal lpException As Exception)
      MyBase.New(String.Format("ID:{0}:{1}", lpExportArgs.ObjectId, lpException.Message), lpException)
      Try
        mstrObjectId = lpExportArgs.ObjectId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportObjectEventArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(String.Format("ID:{0}:{1}", lpExportArgs.ObjectId, lpException.Message), lpException, lpWorker)
      Try
        mstrObjectId = lpExportArgs.ObjectId
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpSourceObjectId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpMessage, lpException)
      mstrObjectId = lpSourceObjectId
    End Sub

    Public Sub New(ByVal lpSourceObjectId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpMessage, lpException, lpWorker)
      mstrObjectId = lpSourceObjectId
    End Sub

#End Region


#Region "Private Methods"

    Protected Friend Overrides Function DebuggerIdentifier() As String
      Try
        If ObjectId.Length > 0 Then
          Return String.Format("ObjectId={0}; {1}", ObjectId, MyBase.DebuggerIdentifier)
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