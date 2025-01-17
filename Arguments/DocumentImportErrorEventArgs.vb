'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel

Namespace Arguments

  ''' <summary>
  ''' Contains all the parameters for the ExportError Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentImportErrorEventArgs
    Inherits DocumentErrorEventArgs

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
                   ByVal lpMessage As String,
                   ByVal lpException As Exception)
      MyBase.New(lpDocumentId, lpMessage, lpException)
    End Sub

    Public Sub New(ByVal lpDocumentId As String,
                   ByVal lpMessage As String,
                   ByVal lpException As Exception,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpDocumentId, lpMessage, lpException, lpWorker)
    End Sub

#End Region

  End Class

End Namespace