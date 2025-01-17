'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations

  ''' <summary>
  ''' Transforms a related document
  ''' </summary>
  ''' <remarks></remarks>
  Public Class TransformAllRelatedDocumentsAction
    Inherits Action

#Region "Class Constants"

    Private Const ACTION_NAME As String = "TransformAllRelatedDocuments"

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try
        If TypeOf Transformation.Target Is Folder Then
          Throw New InvalidTransformationTargetException()
        End If

        If Transformation.Document IsNot Nothing Then
          If Transformation.Document.Relationships IsNot Nothing Then
            For Each lobjRelationship As Relationship In Transformation.Document.Relationships
              'lobjRelationship.RelatedDocument.Transform(Me.Transformation)
              TransformAllDocuments(Me.Transformation)
            Next
            Return New ActionResult(Me, True)
          Else
            Return New ActionResult(Me, False, String.Format("There are no relationships in document {0}.", Transformation.Document.ID))
          End If
        Else
          Return New ActionResult(Me, False, "The associated document reference is not set.")
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        lpErrorMessage &= "TransformAllRelatedDocuments Failed; " & Helper.FormatCallStack(ex)
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        ' Do nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Function TransformAllDocuments(ByVal lpTransformation As Transformation) As ActionResult

#If NET8_0_OR_GREATER Then
      ArgumentNullException.ThrowIfNull(lpTransformation)
#Else
        If lpTransformation Is Nothing Then
          Throw New ArgumentNullException
        End If
#End If

      Try
        Throw New NotImplementedException
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace