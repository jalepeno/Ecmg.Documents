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
  ''' <summary>
  ''' Contains all the parameters for the AnnotationExported Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class AnnotationExportedEventArgs
    Inherits BackgroundEventArgs

#Region "Class Variables"

    Private mobjAnnotation As Annotation

#End Region

#Region "Public Properties"

    Public Property Annotation() As Annotation
      Get
        Return mobjAnnotation
      End Get
      Set(ByVal value As Annotation)
        mobjAnnotation = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpAnnotation As Annotation)
      Me.New(lpAnnotation, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpTime As DateTime)
      Me.New(lpAnnotation, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpExportAnnotationArgs As ExportAnnotationEventArgs)
      Me.New(lpExportAnnotationArgs.Annotation, Now, lpExportAnnotationArgs.ErrorMessage, lpExportAnnotationArgs.Worker)
    End Sub

    Public Sub New(ByVal lpExportAnnotationArgs As ExportAnnotationEventArgs, ByVal lpTime As DateTime)
      Me.New(lpExportAnnotationArgs.Annotation, lpTime, lpExportAnnotationArgs.ErrorMessage, lpExportAnnotationArgs.Worker)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpTime As DateTime, ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpTime, lpWorker)
      Annotation = lpAnnotation

    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation, ByVal lpTime As DateTime, ByVal lpErrorMessage As String, ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpTime, lpErrorMessage, lpWorker)
      Annotation = lpAnnotation

    End Sub

#End Region

  End Class
End Namespace