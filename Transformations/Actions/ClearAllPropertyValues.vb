'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ClearAllValues.vb
'   Description :  [type_description_here]
'   Created     :  5/20/2013 9:20:31 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

#End Region

Namespace Transformations

  Public Class ClearAllPropertyValues
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ClearAllPropertyValues"
    Friend Const PARAM_VERSION_INDEX As String = "VersionIndex"

#End Region

#Region "Class Variables"

    Private mintVersionIndex As Integer = Transformation.TRANSFORM_ALL_VERSIONS

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property VersionIndex() As Integer
      Get
        Return mintVersionIndex
      End Get
      Set(ByVal Value As Integer)
        mintVersionIndex = Value
      End Set

    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.ClearAllValues)
    End Sub

    Public Sub New(ByVal lpPropertyName As String)
      MyBase.New(lpPropertyName, ActionType.ClearAllValues)
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      ' Clear all property values
      Try

        If TypeOf Transformation.Target Is Document Then
          If Transformation.Document.PropertyExists(PropertyScope, PropertyName, False) Then
            Transformation.Document.ClearPropertyValue(PropertyScope,
                                              PropertyName,
                                              VersionIndex)

            Return New ActionResult(Me, True)
          Else
            lpErrorMessage = String.Format("Unable to clear values for property '{0}', the property does not exist in the specified property scope of '{1}'.",
                                           PropertyName, PropertyScope.ToString)
            Return New ActionResult(Me, False, lpErrorMessage)
          End If
        ElseIf TypeOf Transformation.Target Is Folder Then
          If Transformation.Folder.PropertyExists(PropertyName, False) Then
            Transformation.Folder.ClearPropertyValue(PropertyName)

            Return New ActionResult(Me, True)
          Else
            lpErrorMessage = String.Format("Unable to clear values for property '{0}', the property does not exist in the specified property scope of '{1}'.",
                                           PropertyName, PropertyScope.ToString)
            Return New ActionResult(Me, False, lpErrorMessage)
          End If
        Else
          Throw New InvalidTransformationTargetException
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Execute", Me.GetType.Name))
        If lpErrorMessage.Length = 0 Then
          lpErrorMessage = ex.Message
        End If
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = MyBase.GetDefaultParameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_VERSION_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_VERSION_INDEX, Transformation.TRANSFORM_ALL_VERSIONS,
            "Specifies which version(s) this action applies to.  -1 means all versions."))
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
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace