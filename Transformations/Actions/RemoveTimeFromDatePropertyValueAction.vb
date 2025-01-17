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

Namespace Transformations
  ''' <summary>
  ''' Action used to remove the time portion from the value of a Date property
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  Public Class RemoveTimeFromDatePropertyValueAction
    Inherits ChangePropertyValue

#Region "Class Constants"

    Private Const ACTION_NAME As String = "RemoveTimeFromDatePropertyValue"
    Friend Const PARAM_VERSION_INDEX As String = "VersionIndex"

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.ChangePropertyValue)
      MyBase.SourceType = ValueSource.Literal

    End Sub

    Public Sub New(ByVal lpPropertyName As String,
                   Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty,
                   Optional ByVal lpVersionIndex As Integer = Transformation.TRANSFORM_ALL_VERSIONS)

      MyBase.New(lpPropertyName, ActionType.ChangePropertyValue)

      PropertyScope = lpPropertyScope
      VersionIndex = lpVersionIndex
      MyBase.SourceType = ValueSource.Literal

    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult

      Try

        '  Get the property in it's current state
        Dim lobjProperty As ECMProperty = GetTargetProperty()

        If lobjProperty Is Nothing Then
          Return New ActionResult(Me, False, String.Format("No property by the name '{0}' was found.", Me.PropertyName))
        End If

        lpErrorMessage = String.Empty
        If TypeOf Transformation.Target Is Document Then
          RemoveTime(lobjProperty, Transformation.Document, PropertyScope, VersionIndex, lpErrorMessage)
        ElseIf TypeOf Transformation.Target Is Folder Then
          RemoveTime(lobjProperty, Transformation.Folder, lpErrorMessage)
        End If

        If lpErrorMessage.Length = 0 Then
          Return New ActionResult(Me, True)
        Else
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return New ActionResult(Me, False, ex.Message)
      End Try

    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_PROPERTY_NAME) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PROPERTY_NAME, String.Empty,
            "Specifies the name of the target property to modify."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_SCOPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PROPERTY_SCOPE, PropertyScope.VersionProperty, GetType(PropertyScope),
            "Specifies the scope of the target property."))
        End If

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
        MyBase.InitializeParameterValues()
        Me.PropertyName = GetStringParameterValue(PARAM_PROPERTY_NAME, String.Empty)
        Me.PropertyScope = GetEnumParameterValue(PARAM_PROPERTY_SCOPE, GetType(PropertyScope), PropertyScope.VersionProperty)
        Me.VersionIndex = GetParameterValue(PARAM_VERSION_INDEX, Transformation.TRANSFORM_ALL_VERSIONS)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Shared Methods"

    Friend Shared Sub RemoveTime(ByVal lpProperty As ECMProperty,
                                           ByVal lpDocument As Document,
                                           ByVal lpPropertyScope As Integer,
                                           ByVal lpVersionIndex As Integer,
                                           ByRef lpErrorMessage As String)

      Dim lstrErrorMessage As String = String.Empty

      Try

        If ChangePropertyValue.ValidateDateProperty(lpProperty, lpDocument, lpPropertyScope, lpErrorMessage) = True Then

          ' Create a variable to hold the date value in a DateTime object
          Dim ldatValue As DateTime = CType(lpProperty.Value, DateTime)

          ' Strip off the time
          ldatValue = ldatValue.Date

          '  Change the value
          lpDocument.ChangePropertyValue(lpPropertyScope, lpProperty.Name, ldatValue, lpVersionIndex)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lstrErrorMessage.Length > 0 Then
          If lpErrorMessage Is Nothing OrElse lpErrorMessage.Length = 0 Then
            lpErrorMessage = lstrErrorMessage
          Else
            lpErrorMessage = String.Format("; {0}", lstrErrorMessage)
          End If
        End If
      End Try
    End Sub

    Friend Shared Sub RemoveTime(ByVal lpProperty As ECMProperty,
                                       ByVal lpFolder As Folder,
                                       ByRef lpErrorMessage As String)

      Dim lstrErrorMessage As String = String.Empty

      Try

        If ChangePropertyValue.ValidateDateProperty(lpProperty, lpFolder, lpErrorMessage) = True Then

          ' Create a variable to hold the date value in a DateTime object
          Dim ldatValue As DateTime = CType(lpProperty.Value, DateTime)

          ' Strip off the time
          ldatValue = ldatValue.Date

          '  Change the value
          lpFolder.ChangePropertyValue(lpProperty.Name, ldatValue)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      Finally
        If lstrErrorMessage.Length > 0 Then
          If lpErrorMessage Is Nothing OrElse lpErrorMessage.Length = 0 Then
            lpErrorMessage = lstrErrorMessage
          Else
            lpErrorMessage = String.Format("; {0}", lstrErrorMessage)
          End If
        End If
      End Try
    End Sub

#End Region

  End Class

End Namespace