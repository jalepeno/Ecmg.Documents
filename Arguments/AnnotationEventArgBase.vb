'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.Annotations

#End Region

Namespace Arguments

  Public MustInherit Class AnnotationEventArgBase
    Inherits BackgroundEventArgs

#Region "Class Variables"

    Private mstrID As String
    Private mobjAnnotation As Annotation

#End Region

#Region "Public Properties"

    Public ReadOnly Property ID() As String
      Get
        Return mstrID
      End Get
    End Property

    Public ReadOnly Property Annotation() As Annotation
      Get
        Return mobjAnnotation
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpId As String)
      MyBase.New(Now)
      mstrID = lpId
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation)
      Me.New(lpAnnotation, lpAnnotation.ID, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpNewId As String)
      Me.New(lpAnnotation, lpNewId, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpTime As DateTime)
      Me.New(lpAnnotation, lpAnnotation.ID, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpNewId As String, ByVal lpTime As DateTime)
      Me.New(lpAnnotation, lpNewId, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpExportAnnotationArgs As ExportAnnotationEventArgs, ByVal lpNewId As String)
      Me.New(lpExportAnnotationArgs, lpNewId, Now)
    End Sub

    Public Sub New(ByVal lpExportAnnotationArgs As ExportAnnotationEventArgs, ByVal lpNewId As String, ByVal lpTime As DateTime)
      Me.New(lpExportAnnotationArgs.Annotation, lpNewId, lpTime, lpExportAnnotationArgs.Worker)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpNewId As String, ByVal lpTime As DateTime, ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpTime, lpWorker)
      mstrID = lpNewId
      If lpAnnotation Is Nothing Then
        Throw New ArgumentException("lpAnnotation")
      End If
      mobjAnnotation = lpAnnotation
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpNewId As String, ByVal lpTime As DateTime, ByRef lpErrorMessage As String, ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpTime, lpErrorMessage, lpWorker)
      mstrID = lpNewId
      If lpAnnotation Is Nothing Then
        Throw New ArgumentException("lpAnnotation")
      End If
      mobjAnnotation = lpAnnotation
    End Sub

#End Region

  End Class

End Namespace