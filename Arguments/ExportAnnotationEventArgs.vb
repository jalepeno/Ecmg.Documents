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

  Public Class ExportAnnotationEventArgs
    Inherits AnnotationEventArgBase

#Region "Constructors"

    Public Sub New(ByVal lpId As String)
      MyBase.New(lpId)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation,
                   ByVal lpId As String)
      MyBase.New(lpAnnotation, lpId)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime)
      MyBase.New(lpAnnotation, lpId, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpExportAnnotationArgs As ExportAnnotationEventArgs,
                   ByVal lpId As String)
      MyBase.New(lpExportAnnotationArgs, lpId)
    End Sub

    Public Sub New(ByVal lpExportAnnotationArgs As ExportAnnotationEventArgs,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime)
      MyBase.New(lpExportAnnotationArgs.Annotation, lpId, lpTime)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpAnnotation, lpId, lpTime, lpWorker)
    End Sub

    Public Sub New(ByVal lpAnnotation As Annotation,
                   ByVal lpId As String,
                   ByVal lpTime As DateTime,
                   ByRef lpErrorMessage As String,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpAnnotation, lpId, lpTime, lpErrorMessage, lpWorker)
    End Sub

#End Region

  End Class

End Namespace