'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ContentFileHasExtensionTestAction.vb
'   Description :  [type_description_here]
'   Created     :  1/2/2013 2:00:03 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Transformations

  Public Class ContentFileHasExtensionTestAction
    Inherits Action
    Implements IDecisionAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ContentFileHasExtensionTest"
    Friend Const PARAM_VERSION_INDEX As String = "VersionIndex"
    Friend Const PARAM_CONTENT_INDEX As String = "ContentIndex"

#End Region

#Region "Class Variables"

    Private mintVersionIndex As Integer = Transformation.TRANSFORM_ALL_VERSIONS
    Private mintWorkingVersionIndex As Integer
    Private mintContentIndex As Integer = 0
    Private mblnEvaluation As Nullable(Of Boolean)

#End Region

#Region "Public Properties"

    Public Property VersionIndex() As Integer
      Get
        Return mintVersionIndex
      End Get
      Set(ByVal Value As Integer)
        mintVersionIndex = Value
      End Set
    End Property

    Public Property ContentIndex() As Integer
      Get
        Return mintContentIndex
      End Get
      Set(ByVal Value As Integer)
        mintContentIndex = Value
      End Set
    End Property

    Private ReadOnly Property WorkingVersionIndex As Integer
      Get
        Return mintWorkingVersionIndex
      End Get
    End Property

    Public ReadOnly Property Evaluation As Boolean Implements IDecisionAction.Evaluation
      Get
        Try
          If mblnEvaluation.HasValue = False Then
            mblnEvaluation = EvaluateContentForExtension()
          End If

          If mblnEvaluation.HasValue Then
            Return mblnEvaluation
          Else
            Return False
          End If

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property FalseActions As Actions Implements IDecisionAction.FalseActions

    Public ReadOnly Property RunActions As Actions Implements IDecisionAction.RunActions
      Get
        Try
          If Evaluation = True Then
            Return TrueActions
          Else
            If mblnEvaluation.HasValue Then
              Return FalseActions
            Else
              Return New Actions
            End If
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property TrueActions As Actions Implements IDecisionAction.TrueActions

#End Region

#Region "Constructors"

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try

        Dim lobjActionResult As ActionResult = Nothing
        Dim lstrDetails As String = String.Empty
        ' Dim lstrSourceDocumentClass As String = Transformation.Document.DocumentClass
        If Transformation.Document Is Nothing Then
          Throw New Exceptions.TransformationDocumentReferenceNotSetException(Transformation)
        End If

        If VersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
          mintWorkingVersionIndex = Transformation.Document.LatestVersion.ID
        Else
          mintWorkingVersionIndex = VersionIndex
        End If

        If Not Transformation.Document.Versions.ItemExists(WorkingVersionIndex) Then
          Throw New Exceptions.InvalidVersionSpecificationException(Transformation.Document,
            String.Format("VersionIndex '{0}' is not valid for document.", WorkingVersionIndex))
        End If

        If Transformation.Document.Versions(WorkingVersionIndex).HasContent = False Then
          Throw New Exceptions.DocumentHasNoContentException(Transformation.Document,
            String.Format("Unable to test content for extension, document '{0}' version '{1}' has no content.",
                          Transformation.Document.ID, WorkingVersionIndex))
        End If

        If Transformation.Document.Versions(WorkingVersionIndex).Contents(ContentIndex) Is Nothing Then
          Throw New Exceptions.ItemDoesNotExistException(ContentIndex, String.Format("Unable to test content for extension, document '{0}' version '{1}' does not have a content element with index '{2}'.",
                          Transformation.Document.ID, WorkingVersionIndex, ContentIndex))
        End If

        mblnEvaluation = EvaluateContentForExtension()

        If Evaluation = True Then
          For Each lobjTrueAction As Action In TrueActions
            lobjTrueAction.Transformation = Me.Transformation
            lobjTrueAction.Execute(lpErrorMessage)
          Next
        ElseIf Evaluation = False Then
          For Each lobjFalseAction As Action In FalseActions
            lobjFalseAction.Transformation = Me.Transformation
            lobjFalseAction.Execute(lpErrorMessage)
          Next
        End If

        lobjActionResult = New ActionResult(Me, True, lstrDetails)

        Return lobjActionResult

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        If lpErrorMessage.Length = 0 Then
          lpErrorMessage = ex.Message
        End If
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try
    End Function

    Public Sub CombineActions() Implements IDecisionAction.CombineActions
      Try

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = MyBase.GetDefaultParameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_VERSION_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_VERSION_INDEX, Transformation.TRANSFORM_ALL_VERSIONS,
            "Specifies the specific version of the document to test. The index is zero based, meaning the first version is specified with a value of zero. The value -1 may also be specified to test the latest version."))
        End If

        If lobjParameters.Contains(PARAM_CONTENT_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_CONTENT_INDEX, Transformation.TRANSFORM_ALL_VERSIONS,
            "Specifies the specific content element of the document to test. The index is zero based, meaning the first content element is specified with a value of zero."))
        End If

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        Me.VersionIndex = GetParameterValue(PARAM_VERSION_INDEX, Transformation.TRANSFORM_ALL_VERSIONS)
        Me.ContentIndex = GetParameterValue(PARAM_CONTENT_INDEX, Transformation.TRANSFORM_ALL_VERSIONS)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Private Function EvaluateContentForExtension() As Nullable(Of Boolean)
      Try
        If Transformation.Document Is Nothing Then
          Return False
        Else
          Return Transformation.Document.Versions(WorkingVersionIndex).Contents(ContentIndex).HasExtension
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace