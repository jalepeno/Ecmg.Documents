'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel

#End Region


Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters for the ExportError Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class AnnotationExportErrorEventArgs
    Inherits ErrorEventArgs

#Region "Class Variables"

    Private mstrAnnotationId As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property AnnotationId() As String
      Get
        Return mstrAnnotationId
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

    Public Sub New(ByVal lpExportArgs As ExportAnnotationEventArgs,
                   ByVal lpException As Exception)
      MyBase.New(lpExportArgs.ErrorMessage, lpException, lpExportArgs.Worker)
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportAnnotationEventArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpExportArgs.ErrorMessage, lpException, lpWorker)
    End Sub

    Public Sub New(ByVal lpAnnotationId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpMessage, lpException)
      mstrAnnotationId = lpAnnotationId
    End Sub

    Public Sub New(ByVal lpAnnotationId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpMessage, lpException, lpWorker)
      mstrAnnotationId = lpAnnotationId
    End Sub

#End Region

  End Class

End Namespace