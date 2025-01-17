'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Annotations

  ''' <summary>
  ''' A collection wrapper for <see cref="Annotation"/> to provide proper XML serialization output.
  ''' </summary>
  <Serializable()>
  <DebuggerDisplay("Annotation Collection")>
  Public Class Annotations
    Inherits CCollection(Of Annotation)

    Public Function AnnotatedContentIds() As List(Of String)
      Dim result As New List(Of String)
      Try
        For Each member As Annotation In Me.Items
          Dim annotatedId As String = member.AnnotatedContentId
          If Not result.Contains(annotatedId) Then
            result.Add(annotatedId)
          End If
        Next

        Return result

      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ContentAnnotations(ByVal contentId As String) As List(Of Annotation)
      Dim result As New List(Of Annotation)
      Try
        For Each member As Annotation In Me.Items
          Dim annotatedId As String = member.AnnotatedContentId
          If String.Equals(contentId, annotatedId, StringComparison.InvariantCultureIgnoreCase) Then
            result.Add(member)
          End If
        Next

        Return result

      Catch ex As System.Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Shadows ReadOnly Property Items As IList(Of Annotation)
      Get
        Return MyBase.Items
      End Get
    End Property

  End Class

End Namespace