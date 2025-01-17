'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel
Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentExported Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentExportedEventArgs
    Inherits BackgroundEventArgs

#Region "Class Variables"

    Private mobjDocument As Document

#End Region

#Region "Public Properties"

    Public Property Document() As Document
      Get
        Return mobjDocument
      End Get
      Set(ByVal value As Document)
        mobjDocument = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document)
      Me.New(lpDocument, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpTime As DateTime)
      Me.New(lpDocument, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpExportDocumentArgs As ExportDocumentEventArgs)
      Me.New(lpExportDocumentArgs.Document, Now, lpExportDocumentArgs.ErrorMessage, lpExportDocumentArgs.Worker)
    End Sub

    Public Sub New(ByVal lpExportDocumentArgs As ExportDocumentEventArgs, ByVal lpTime As DateTime)
      Me.New(lpExportDocumentArgs.Document, lpTime, lpExportDocumentArgs.ErrorMessage, lpExportDocumentArgs.Worker)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpTime As DateTime, ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpTime, lpWorker)
      Document = lpDocument

    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpTime As DateTime, ByVal lpErrorMessage As String, ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpTime, lpErrorMessage, lpWorker)
      Document = lpDocument

    End Sub

#End Region

  End Class
End Namespace