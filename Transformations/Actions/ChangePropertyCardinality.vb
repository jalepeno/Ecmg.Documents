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
  ''' Changes the cardinality of a property to either single-valued or multi-valued.
  ''' </summary>
  ''' <remarks>If the existing cardinality of the property is the same as the cardinality specified in the change the action will return false with an explanation.</remarks>
  <Serializable()>
  Public Class ChangePropertyCardinality
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ChangePropertyCardinality"
    Friend Const PARAM_CARDINALITY As String = "Cardinality"
    Friend Const PARAM_MULTI_VALUED_PROPERTY_DELIMITER As String = "MultiValuedPropertyDelimiter"

#End Region

#Region "Class Variables"

    Private menuCardinality As Cardinality
    Private mstrMultiValuedPropertyDelimiter As String = ","

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property Cardinality() As Cardinality
      Get
        Return menuCardinality
      End Get
      Set(ByVal value As Cardinality)
        menuCardinality = value
      End Set
    End Property

    Public Property MultiValuedPropertyDelimiter() As String
      Get
        Return mstrMultiValuedPropertyDelimiter
      End Get
      Set(ByVal value As String)
        mstrMultiValuedPropertyDelimiter = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.ChangePropertyCardinality)
    End Sub

    Public Sub New(ByVal lpPropertyName As String,
                   ByVal lpPropertyScope As Core.PropertyScope,
                   ByVal lpCardinality As Cardinality,
                   ByVal lpMultiValuedPropertyDelimiter As String)

      MyBase.New(lpPropertyName, lpPropertyScope)

      Cardinality = lpCardinality
      MultiValuedPropertyDelimiter = lpMultiValuedPropertyDelimiter

    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult

      ' Change the cardinality of the property
      Dim lobjActionResult As ActionResult = Nothing
      Dim lobjProperty As ECMProperty

      Try

        If TypeOf Transformation.Target Is Document Then
          ' Make sure the property is defined in the document
          If Transformation.Document.PropertyExists(Me.PropertyScope, Me.PropertyName) = False Then
            lpErrorMessage &= String.Format("The target property '{0}' of scope '{1}' is not defined in the document.",
                                            Me.PropertyName, Me.PropertyScope.ToString)
            Return New ActionResult(Me, False, lpErrorMessage)
          End If

          Select Case Me.PropertyScope
            Case PropertyScope.DocumentProperty

              ' Get a reference to the property
              lobjProperty = Transformation.Document.Properties(Me.PropertyName)

              ' Set the cardinality of the property
              lobjActionResult = ChangePropertyCardinality(lobjProperty)
              If lobjActionResult.Success = True Then
                Transformation.Document.Properties.Replace(Me.PropertyName, lobjProperty)
              Else
                lpErrorMessage = lobjActionResult.Details
              End If
              Return lobjActionResult

            Case PropertyScope.VersionProperty

              ' If we are going to change the cardinality of a version property we need to do it for each version of the document
              For Each lobjVersion As Version In Transformation.Document.Versions
                ' Get a reference to the property
                lobjProperty = lobjVersion.Properties(Me.PropertyName)
                lobjActionResult = ChangePropertyCardinality(lobjProperty)
                If lobjActionResult.Success = True Then
                  lobjVersion.Properties.Replace(Me.PropertyName, lobjProperty)
                Else
                  ' If this version failed then we will not try any others
                  lpErrorMessage = lobjActionResult.Details
                  Return lobjActionResult
                End If
              Next

              ' If we are still here then we succeeded with all of the versions
              Return lobjActionResult

            Case PropertyScope.BothDocumentAndVersionProperties

              ' First change it at the document level
              ' Get a reference to the property
              lobjProperty = Transformation.Document.Properties(Me.PropertyName)

              ' Set the cardinality of the property
              lobjActionResult = ChangePropertyCardinality(lobjProperty)
              If lobjActionResult.Success = True Then
                Transformation.Document.Properties.Replace(Me.PropertyName, lobjProperty)
              Else
                ' If this version failed then we will not try any others
                lpErrorMessage = lobjActionResult.Details
                Return lobjActionResult
              End If

              ' Now change it at the version level
              ' If we are going to change the cardinality of a version property we need to do it for each version of the document
              For Each lobjVersion As Version In Transformation.Document.Versions
                ' Get a reference to the property
                lobjProperty = lobjVersion.Properties(Me.PropertyName)
                lobjActionResult = ChangePropertyCardinality(lobjProperty)
                If lobjActionResult.Success = True Then
                  lobjVersion.Properties.Replace(Me.PropertyName, lobjProperty)
                Else
                  ' If this version failed then we will not try any others
                  lpErrorMessage = lobjActionResult.Details
                  Return lobjActionResult
                End If
              Next

              ' If we are still here then we succeeded with all of the versions
              Return lobjActionResult

            Case Else
              ' If we are here it is because we were passed an invalid scope
              lpErrorMessage &= String.Format("The scope '{0}' is not defined.",
                                              Me.PropertyName, Me.PropertyScope)
              Return New ActionResult(Me, False, lpErrorMessage)
          End Select

        ElseIf TypeOf Transformation.Target Is Folder Then
          ' Make sure the property is defined in the folder
          If Transformation.Folder.PropertyExists(Me.PropertyName) = False Then
            lpErrorMessage &= String.Format("The target property '{0}' of scope '{1}' is not defined in the folder.",
                                            Me.PropertyName, Me.PropertyScope.ToString)
            Return New ActionResult(Me, False, lpErrorMessage)
          End If

          ' Get a reference to the property
          lobjProperty = Transformation.Folder.Properties(Me.PropertyName)

          ' Set the cardinality of the property
          lobjActionResult = ChangePropertyCardinality(lobjProperty)
          If lobjActionResult.Success = True Then
            Transformation.Folder.Properties.Replace(Me.PropertyName, lobjProperty)
          Else
            lpErrorMessage = lobjActionResult.Details
          End If

          Return lobjActionResult

        Else
          Throw New InvalidTransformationTargetException

        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Execute", Me.GetType.Name))
        lobjActionResult = New ActionResult(Me, False, Helper.FormatCallStack(ex))
        Return lobjActionResult
      End Try

    End Function

    ''' <summary>
    ''' Change the cardinality of a single property
    ''' </summary>
    ''' <param name="lpProperty">The property to update</param>
    ''' <returns>An ActionResult object specifying the result</returns>
    ''' <remarks></remarks>
    Private Function ChangePropertyCardinality(ByRef lpProperty As ECMProperty) As ActionResult

      Dim lobjActionResult As ActionResult = Nothing

      Try

        ' Make sure the current cardinality is not the same as the target cardinality
        If lpProperty.Cardinality = Me.Cardinality Then
          lobjActionResult = New ActionResult(Me, False,
                                              String.Format("The property '{0}' already has a cardinality of '{1}', the cardinality change is not necessary.",
                                                            lpProperty.Name, lpProperty.Cardinality.ToString))
          Return lobjActionResult
        End If

        Select Case Me.Cardinality
          Case Core.Cardinality.ecmMultiValued
            ' If the property is of the newer style we can convert by calling ToMultiValue
            If TypeOf (lpProperty) Is SingletonProperty Then
              lpProperty = DirectCast(lpProperty, SingletonProperty).ToMultiValue(Me.MultiValuedPropertyDelimiter)
            Else
              ' This is an older style property that is of the base type ECMProperty
              ' We will have to do this the original way.


              Dim lpPropertySavedValue As New Value(lpProperty.Value.ToString)
              lpProperty.Value = Nothing
              lpProperty.Cardinality = Me.Cardinality

              ' Try to move the value to the values collection
              If lpPropertySavedValue IsNot Nothing Then
                lpProperty.Values.Clear()
                Dim lobjValue As New Value
                Dim lstrValues() As String = lpPropertySavedValue.ToString.Split(MultiValuedPropertyDelimiter)
                For Each lstrValue As String In lstrValues
                  If lstrValue.Length > 0 Then
                    lobjValue = New Value(lstrValue)
                    lpProperty.Values.Add(lobjValue)
                  End If
                Next
              End If
            End If

          Case Core.Cardinality.ecmSingleValued
            ' If the property is of the newer style we can convert by calling ToMultiValue
            If TypeOf (lpProperty) Is MultiValueProperty Then
              lpProperty = DirectCast(lpProperty, MultiValueProperty).ToSingleton(MultiValuedPropertyDelimiter)
            Else
              ' This is an older style property that is of the base type ECMProperty
              ' We will have to do this the original way.

              ' Try to move the values collection into a single value
              lpProperty.Value = Nothing
              lpProperty.Cardinality = Me.Cardinality
              Dim lstrDelimitedProperty As String = String.Empty
              For Each lobjValue As Object In lpProperty.Values
                If TypeOf (lobjValue) Is Value Then
                  lstrDelimitedProperty += lobjValue.Value.ToString & MultiValuedPropertyDelimiter
                Else
                  lstrDelimitedProperty += lobjValue.ToString & MultiValuedPropertyDelimiter
                End If
              Next

              If lstrDelimitedProperty.EndsWith(MultiValuedPropertyDelimiter) Then
                lstrDelimitedProperty = lstrDelimitedProperty.Remove(lstrDelimitedProperty.Length - MultiValuedPropertyDelimiter.Length)
              End If

              lpProperty.Value = lstrDelimitedProperty
              lpProperty.Values = New Values
            End If
        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, "ChangePropertyCardinality::ChangePropertyCardinality")
        lobjActionResult = New ActionResult(Me, False, ex.Message)
        Return lobjActionResult
      End Try

      lobjActionResult = New ActionResult(Me, True,
                                          String.Format("Successfully changed the cardinality of '{0}' to '{1}'.",
                                                        lpProperty.Name, lpProperty.Cardinality.ToString))
      Return lobjActionResult

    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        lobjParameters = MyBase.GetDefaultParameters

        If lobjParameters.Contains(PARAM_CARDINALITY) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_CARDINALITY, Cardinality.ecmSingleValued, GetType(Cardinality),
            "Specifies the new cardinality of the property."))
        End If

        If lobjParameters.Contains(PARAM_MULTI_VALUED_PROPERTY_DELIMITER) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_MULTI_VALUED_PROPERTY_DELIMITER, ",",
            "Specifies delimiter to use for multi-valued properties."))
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
        Me.Cardinality = GetEnumParameterValue(PARAM_CARDINALITY, GetType(Cardinality), Cardinality.ecmSingleValued)
        Me.MultiValuedPropertyDelimiter = GetStringParameterValue(PARAM_MULTI_VALUED_PROPERTY_DELIMITER, ",")
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace