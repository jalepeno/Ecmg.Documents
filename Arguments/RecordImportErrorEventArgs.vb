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

  Public Class RecordImportErrorEventArgs
    Inherits DocumentImportErrorEventArgs

#Region "Public Properties"

    Public Property Record As Object

#End Region

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

    Public Sub New(ByVal lpExportArgs As ImportDocumentArgs,
                   ByVal lpException As Exception)
      MyBase.New(lpExportArgs, lpException)
    End Sub

    Public Sub New(ByVal lpExportArgs As ImportDocumentArgs,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpExportArgs, lpException, lpWorker)
    End Sub

    Public Sub New(ByVal lpDocumentId As String,
                   ByVal lpRecord As Object,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpDocumentId, lpMessage, lpException)
      Record = lpRecord
    End Sub

    Public Sub New(ByVal lpDocumentId As String,
                   ByVal lpRecord As Object,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpDocumentId, lpMessage, lpException, lpWorker)
      Record = lpRecord
    End Sub

#End Region

  End Class

End Namespace