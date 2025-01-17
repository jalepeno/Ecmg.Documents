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
  ''' Contains all the parameters for the DocumentClassImported Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentClassImportedEventArgs
    Inherits BackgroundEventArgs

#Region "Class Variables"

    Private mstrNewID As String
    Private mobjDocumentClass As DocumentClass

#End Region

#Region "Public Properties"

    Public ReadOnly Property NewID() As String
      Get
        Return mstrNewID
      End Get
    End Property

    Public ReadOnly Property NewDocumentClass() As DocumentClass
      Get
        Return mobjDocumentClass
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocumentClass As DocumentClass, ByVal lpNewId As String)
      Me.New(lpDocumentClass, lpNewId, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpDocumentClass As DocumentClass, ByVal lpNewId As String, ByVal lpTime As DateTime)
      Me.New(lpDocumentClass, lpNewId, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpImportDocumentClassArgs As ImportDocumentClassEventArgs, ByVal lpNewId As String)
      Me.New(lpImportDocumentClassArgs, lpNewId, Now)
    End Sub

    Public Sub New(ByVal lpImportDocumentClassArgs As ImportDocumentClassEventArgs, ByVal lpNewId As String, ByVal lpTime As DateTime)
      Me.New(lpImportDocumentClassArgs.DocumentClass, lpNewId, lpTime, lpImportDocumentClassArgs.Worker)
    End Sub

    Public Sub New(ByVal lpDocumentClass As DocumentClass, ByVal lpNewId As String, ByVal lpTime As DateTime, ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpTime, lpWorker)
      mstrNewID = lpNewId
      If lpDocumentClass Is Nothing Then
        Throw New ArgumentException("lpDocumentClass")
      End If
      mobjDocumentClass = lpDocumentClass
    End Sub

    Public Sub New(ByVal lpDocumentClass As DocumentClass, ByVal lpNewId As String, ByVal lpTime As DateTime, ByRef lpErrorMessage As String, ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpTime, lpErrorMessage, lpWorker)
      mstrNewID = lpNewId
      If lpDocumentClass Is Nothing Then
        Throw New ArgumentException("lpDocumentClass")
      End If
      mobjDocumentClass = lpDocumentClass
    End Sub

#End Region

  End Class
End Namespace