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

  Public Class MigrateDocumentErrorEventArgs
    Inherits DocumentExportErrorEventArgs

#Region "Constructors"

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
      MyBase.New(lpExportArgs, lpException)
    End Sub

    Public Sub New(ByVal lpExportArgs As ExportDocumentEventArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpExportArgs, lpException, lpWorker)
    End Sub

    Public Sub New(ByVal lpSourceDocumentId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpSourceDocumentId, lpMessage, lpException)
    End Sub

    Public Sub New(ByVal lpSourceDocumentId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpSourceDocumentId, lpMessage, lpException, lpWorker)
    End Sub

#End Region

  End Class

End Namespace