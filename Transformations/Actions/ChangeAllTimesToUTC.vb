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

  <Serializable()>
  Public Class ChangeAllTimesToUTC
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ChangeAllTimesToUTC"

#End Region

    Public Sub New()
      MyBase.New(ActionType.ChangeAllTimesToUTC)
    End Sub

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_PROPERTY_SCOPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PROPERTY_SCOPE, String.Empty,
            "Specifies the scope of the target properties."))
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
        ' Do nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub


    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      Try

        lpErrorMessage = String.Empty

        If TypeOf Transformation.Target Is Document Then
          '  Get all the date properties at the document level
          For Each lobjProperty As ECMProperty In Transformation.Document.PropertiesByType(PropertyType.ecmDate)
            ChangeDateTimePropertyValueToUTC.ChangeDateValueToUTC(lobjProperty, Transformation.Document,
                                                                    PropertyScope.DocumentProperty,
                                                                     0, lpErrorMessage)
          Next

          '  Get all the date properties at the version level
          Dim lintVersionCounter As Integer = 0
          For Each lobjVersion As Version In Transformation.Document.Versions
            For Each lobjProperty As ECMProperty In lobjVersion.PropertiesByType(PropertyType.ecmDate)
              ChangeDateTimePropertyValueToUTC.ChangeDateValueToUTC(lobjProperty,
                                                                      Transformation.Document,
                                                                      PropertyScope.VersionProperty,
                                                                      lintVersionCounter,
                                                                      lpErrorMessage)
            Next
            lintVersionCounter += 1
          Next
        ElseIf TypeOf Transformation.Target Is Folder Then
          For Each lobjProperty As ECMProperty In Transformation.Folder.PropertiesByType(PropertyType.ecmDate)
            ChangeDateTimePropertyValueToUTC.ChangeDateValueToUTC(lobjProperty, Transformation.Folder,
                                                                    lpErrorMessage)
          Next
        End If


        Return New ActionResult(Me, True)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return New ActionResult(Me, False, ex.Message)
      End Try
    End Function

  End Class

End Namespace