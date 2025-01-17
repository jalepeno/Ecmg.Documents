'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel
Imports Documents.Core

Namespace Arguments
  ''' <summary>Contains all the parameters for the FolderDocumentExported Event</summary>
  Public Class FolderDocumentExportedEventArgs
    Inherits DocumentExportedEventArgs

#Region "Class Variables"

    Private mintDocumentCounter As Integer
    Private mintFolderDocumentCount As Integer

#End Region

#Region "Public Properties"

    Public ReadOnly Property DocumentCounter() As Integer
      Get
        Return mintDocumentCounter
      End Get
    End Property

    Public ReadOnly Property FolderDocumentCount() As Integer
      Get
        Return mintFolderDocumentCount
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document,
                   ByVal lpTime As DateTime,
                   ByVal lpDocumentCounter As Integer,
                   ByVal lpFolderDocumentCount As Integer,
                   ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpDocument, lpTime, lpWorker)
      mintDocumentCounter = lpDocumentCounter
      mintFolderDocumentCount = lpFolderDocumentCount

    End Sub

#End Region

  End Class
End Namespace