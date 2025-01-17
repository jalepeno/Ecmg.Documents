'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core
Imports Documents.Exceptions
Imports Documents.Utilities

Namespace Transformations

  <Serializable()>
  Public Class AddPropertyValue
    Inherits Transformations.ChangePropertyValue

#Region "Class Constants"

    Private Const ACTION_NAME As String = "AddPropertyValue"

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

        If SourceExists() = False Then
          ' We were not able to verify the source property above
          If TypeOf (DataLookup) Is IPropertyLookup Then
            lpErrorMessage = String.Format("Unable to add property value {0}, the source property {1} does not exist.",
                                           PropertyName, CType(DataLookup, IPropertyLookup).SourceProperty.PropertyName)
          Else
            lpErrorMessage = String.Format("Unable to add property value {0}, the source does not exist.", PropertyName)
          End If
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        Dim lobjDestinationProperty As ECMProperty
        If TypeOf Transformation.Target Is Document Then
          lobjDestinationProperty = Transformation.Document.GetProperty(PropertyName, PropertyScope, VersionIndex)
        ElseIf TypeOf Transformation.Target Is Folder Then
          lobjDestinationProperty = Transformation.Folder.GetProperty(PropertyName)
        Else
          Throw New InvalidTransformationTargetException
        End If

        If lobjDestinationProperty Is Nothing AndAlso
          Me.SourceType = ValueSource.DataLookup AndAlso
          DataLookup IsNot Nothing AndAlso
          CType(DataLookup, IPropertyLookup).DestinationProperty.AutoCreate = True Then

          lobjDestinationProperty = CType(DataLookup, IPropertyLookup).DestinationProperty.CreateProperty(Transformation.Document)
          lobjDestinationProperty.Cardinality = Cardinality.ecmMultiValued

        End If

        If lobjDestinationProperty Is Nothing Then
          If Me.SourceType = ValueSource.DataLookup Then
            ' We were unable to get or create the destination
            lpErrorMessage = String.Format("Unable to add value to '{0}', the destination property '{0}' does not exist and AutoCreate is set to false.",
                                             PropertyName)
          Else
            ' We were unable to get or create the destination
            lpErrorMessage = String.Format("Unable to add value to '{0}', the destination property '{0}' does not exist.",
                                             PropertyName)
          End If
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        ' Make sure the destination property is multi-valued
        If lobjDestinationProperty.Cardinality <> Cardinality.ecmMultiValued Then
          lpErrorMessage = String.Format("Unable to perform AddPropertyValue action {0}, the destination property '{1}' is a singlevalued property, this action is only valid for multi-valued destination properties.", Me.Name, PropertyName)
          ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error)
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        Select Case SourceType
          Case ChangePropertyValue.ValueSource.Literal
            Try
              'Transformation.Document.ChangePropertyValue(PropertyScope, _
              '                                            PropertyName, _
              '                                            PropertyValue, _
              '                                            VersionIndex)
              lobjDestinationProperty.AddValue(PropertyValue, AllowDuplicates)

              Return New ActionResult(Me, True)

            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              Return New ActionResult(Me, False, ex.Message)
            End Try

          Case ChangePropertyValue.ValueSource.DataLookup
            'debug.writeline(lobjChangePropertyValueAction.DataMap.SQLStatement(lpEcmDocument))
            Try
              Dim lobjValue As Object
              If Me.VersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
                For lintVersionCounter As Integer = 0 To Transformation.Document.Versions.Count - 1
                  Dim lobjVersion As Version = Transformation.Document.Versions(lintVersionCounter)
                  lobjValue = DataLookup.GetValue(lobjVersion)
                  lobjDestinationProperty = lobjVersion.Properties(PropertyName)
                  'lobjDestinationProperty.Values.Add(lobjValue, AllowDuplicates)
                  lobjDestinationProperty.AddValue(lobjValue, AllowDuplicates)
                Next
              Else
                lobjValue = DataLookup.GetValue(Transformation.Document)
                Transformation.Document.AddPropertyValue(PropertyScope, PropertyName, lobjValue, VersionIndex, AllowDuplicates)
              End If
              Return New ActionResult(Me, True)
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              If ex.Message.StartsWith("No value found for the expression") = True Then
                'Throw New Exception("DataMap found no value for [" & lobjChangePropertyValueAction.PropertyName & "]", ex)
                lpErrorMessage &= "DataMap found no value for [" & PropertyName & "]" & ex.Message
                Return New ActionResult(Me, False, lpErrorMessage)
              Else
                'Throw New Exception("DataLookup Failed", ex)
                lpErrorMessage &= "DataLookup Failed; " & Helper.FormatCallStack(ex)
                Return New ActionResult(Me, False, lpErrorMessage)
              End If
            End Try

        End Select

        Return New ActionResult(Me, True)

      Catch ValueExistsEx As Exceptions.ValueExistsException
        lpErrorMessage = ValueExistsEx.Message
        Return New ActionResult(Me, False, lpErrorMessage)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Overrides Sub InitializeParameterValues()
      Try
        MyBase.InitializeParameterValues()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
