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
  ''' <summary>
  ''' Changes the file name of a content element in the document.
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  Public Class ChangeContentRetrievalName
    Inherits ChangePropertyValue

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ChangeContentRetrievalName"
    Private Const PARAM_CONTENT_ELEMENT_INDEX As String = "ContentElementIndex"

#End Region

#Region "Class Variables"

    Private mintContentElementIndex As Integer

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    Public Property ContentElementIndex() As Integer
      Get
        Return mintContentElementIndex
      End Get
      Set(ByVal value As Integer)
        mintContentElementIndex = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.ChangePropertyValue)
    End Sub

    Public Sub New(ByVal lpPropertyName As String,
                   Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty,
                   Optional ByVal lpVersionIndex As Integer = Transformation.TRANSFORM_ALL_VERSIONS,
                   Optional ByVal lpContentElementIndex As Integer = 0,
                   Optional ByVal lpSourceType As ValueSource = ValueSource.Literal,
                   Optional ByVal lpDataLookup As DataLookup = Nothing,
                   Optional ByVal lpLiteralPropertyValue As Object = Nothing)

      MyBase.New(lpPropertyName, lpPropertyScope, lpVersionIndex, lpSourceType, lpDataLookup, lpLiteralPropertyValue)

      ContentElementIndex = lpContentElementIndex

    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult

      '' Change a property value
      'Dim lobjChangePropertyValueAction As ChangePropertyValue = lobjAction
      Try

        Dim lblnSuccess As Boolean

        ApplicationLogging.WriteLogEntry("Beginning ChangeContentRetrievalName", TraceEventType.Verbose, 4111)

        If TypeOf Transformation.Target Is Folder Then
          Throw New InvalidTransformationTargetException()
        End If

        Dim lobjActionResult As ActionResult = Nothing

        ' Make sure we have a valid Transformation object reference
        If Transformation Is Nothing Then
          lpErrorMessage = "Unable to change content retrieval name, the Transformation property is not initialized."
          ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error, 5391)
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        ' Make sure we have a valid Document object reference
        If Transformation.Document Is Nothing Then
          lpErrorMessage = "Unable to change content retrieval name, the Document property is not initialized."
          ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error, 5492)
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        Select Case SourceType
          Case ChangePropertyValue.ValueSource.Literal
            Try
              Transformation.Document.ChangeContentRetrievalName(PropertyValue, VersionIndex, ContentElementIndex)
              Return New ActionResult(Me, True)

            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              Return New ActionResult(Me, False, ex.Message)
            End Try

          Case ChangePropertyValue.ValueSource.DataLookup
            'debug.writeline(lobjChangePropertyValueAction.DataMap.SQLStatement(lpEcmDocument))
            Try
              Dim lobjValue As Object

              If DataLookup Is Nothing Then
                lpErrorMessage = "Unable to change content retrieval name, the DataLookup property is not initialized."
                ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error, 5401)
                Return New ActionResult(Me, False, lpErrorMessage)
              End If

              If CType(DataLookup, DataParser).SourceProperty Is Nothing Then
                lpErrorMessage = "Unable to change content retrieval name, the SourceProperty property of the DataLookup property is not initialized."
                ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error, 5402)
                Return New ActionResult(Me, False, lpErrorMessage)
              End If

              For Each lobjVersion As Core.Version In Transformation.Document.Versions

                Select Case DataLookup.GetType.Name

                  Case "DataParser"

                    Select Case CType(DataLookup, DataParser).SourceProperty.PropertyScope

                      Case Core.PropertyScope.VersionProperty
                        lobjValue = DataLookup.GetValue(lobjVersion)

                      Case Else
                        lobjValue = DataLookup.GetValue(lobjVersion.GetPrimaryContent)

                    End Select

                  Case "DataList"
                    Select Case CType(DataLookup, DataList).SourceProperty.PropertyScope

                      Case Core.PropertyScope.VersionProperty
                        lobjValue = DataLookup.GetValue(lobjVersion)

                      Case Else
                        lobjValue = DataLookup.GetValue(lobjVersion.GetPrimaryContent)

                    End Select

                  Case Else
                    lobjValue = DataLookup.GetValue(lobjVersion.GetPrimaryContent)
                End Select

                If lobjValue Is Nothing Then
                  lpErrorMessage = String.Format("Unable to change content retrieval name, failed to get the source value for '{0}'.", CType(DataLookup, DataParser).SourceProperty.PropertyName)
                  ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error, 5404)
                  Return New ActionResult(Me, False, lpErrorMessage)
                End If

                If lobjVersion Is Nothing Then
                  lpErrorMessage = "Unable to change content retrieval name, no versions were found."
                  ApplicationLogging.WriteLogEntry(lpErrorMessage, TraceEventType.Error, 5405)
                  Return New ActionResult(Me, False, lpErrorMessage)
                End If

                lblnSuccess = Transformation.Document.ChangeContentRetrievalName(lobjValue.ToString, lobjVersion.ID, ContentElementIndex)

                If lblnSuccess = False Then
                  If ExceptionTracker.LastException IsNot Nothing Then
                    lpErrorMessage = String.Format("Failed to change content retrieval name: '{0}'",
                                                   ExceptionTracker.LastException.Message)
                    Return New ActionResult(Me, False, lpErrorMessage)
                  End If
                End If
              Next

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
            Finally

            End Try

        End Select

        Return New ActionResult(Me, True)

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
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_CONTENT_ELEMENT_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_CONTENT_ELEMENT_INDEX, String.Empty,
            "Specifies which content element to change the retrieval name for."))
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
        Me.ContentElementIndex = GetStringParameterValue(PARAM_CONTENT_ELEMENT_INDEX, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace