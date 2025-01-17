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
  ''' Changes the mime type of a content element in the document.
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()> Public Class ChangeContentMimeType
    Inherits ChangePropertyValue

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ChangeContentMimeType"
    Private Const PARAM_CONTENT_ELEMENT_INDEX As String = "ContentElementIndex"

#End Region

#Region "Class Variables"

    Private mintContentElementIndex As Integer

#End Region

#Region "Public Properties"

    '<XmlAttribute()> _
    'Public Overrides Property Name() As String
    '  Get
    '    If mstrName.Length < 2 Then
    '      mstrName = Me.GetType.Name
    '    End If
    '    Return mstrName
    '  End Get
    '  Set(ByVal value As String)
    '    mstrName = value
    '  End Set
    'End Property

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

    Public Sub New(ByVal lpActionName As String,
                   Optional ByVal lpLiteralPropertyValue As Object = Nothing,
                   Optional ByVal lpVersionIndex As Integer = Transformation.TRANSFORM_ALL_VERSIONS,
                   Optional ByVal lpContentElementIndex As Integer = 0,
                   Optional ByVal lpSourceType As ValueSource = ValueSource.Literal,
                   Optional ByVal lpDataLookup As DataLookup = Nothing)

      MyBase.New(ActionType.ChangePropertyValue)

      With Me
        .PropertyName = "ContentMimeType"
        .mstrName = lpActionName
        .VersionIndex = lpVersionIndex
        .ContentElementIndex = lpContentElementIndex
        .SourceType = lpSourceType
        .DataLookup = lpDataLookup
        .PropertyValue = lpLiteralPropertyValue
      End With

    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult

      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' NOTE: If we are using a MapList we should only change the content mime type if a match 
      ' is found between the current FileExtension and a listed Original value in a ValueMap.
      ' Otherwise we will unintentionally change the mime type to the actual file extension.
      ' This would be a bad thing since it would potentially change an otherwise valid mime type 
      ' to an invalid one.
      ' 
      ' Ernie Bahr
      ' December 9, 2010
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      '' Change a property value
      'Dim lobjChangePropertyValueAction As ChangePropertyValue = lobjAction
      Try

        If TypeOf Transformation.Target Is Folder Then
          Throw New InvalidTransformationTargetException()
        End If

        Dim lobjActionResult As ActionResult = Nothing
        Dim lblnChangeMimeTypeReturn As Nullable(Of Boolean) = Nothing
        Dim lstrChangeMessage As String = String.Empty

        'If SourceExists() = False Then
        '  ' We were not able to verify the source property above
        '  If TypeOf (DataLookup) Is IPropertyLookup Then
        '    lpErrorMessage = String.Format("Unable to change content mime type, the source property {0} does not exist.", _
        '                                   CType(DataLookup, IPropertyLookup).SourceProperty.PropertyName)
        '  Else
        '    lpErrorMessage = "Unable to change content mime type, the source does not exist."
        '  End If
        '  Return New ActionResult(Me, False, lpErrorMessage)
        'End If

        Select Case SourceType

          Case ChangePropertyValue.ValueSource.Literal

            Try
              lblnChangeMimeTypeReturn = Transformation.Document.ChangeContentMimeType(PropertyValue, VersionIndex, ContentElementIndex, lstrChangeMessage)
              If lblnChangeMimeTypeReturn = False Then
                lpErrorMessage = lstrChangeMessage
              End If
              Return New ActionResult(Me, lblnChangeMimeTypeReturn, lpErrorMessage)

            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              Return New ActionResult(Me, False, ex.Message)
            End Try

          Case ChangePropertyValue.ValueSource.DataLookup

            'debug.writeline(lobjChangePropertyValueAction.DataMap.SQLStatement(lpEcmDocument))
            Try

              Dim lobjValue As Object = Nothing

              Dim lobjPropertyLookup As IPropertyLookup
              Dim lobjPropertyLookupInterface As Type = DataLookup.GetType.GetInterface("IPropertyLookup")

              If lobjPropertyLookupInterface IsNot Nothing Then
                lobjPropertyLookup = CType(DataLookup, IPropertyLookup)

                If lobjPropertyLookup.SourceProperty.VersionIndex <> Document.ALL_VERSIONS Then

                  Select Case lobjPropertyLookup.SourceProperty.PropertyScope

                    Case Core.PropertyScope.VersionProperty
                      lobjValue = DataLookup.GetValue(Transformation.Document.Versions(lobjPropertyLookup.SourceProperty.VersionIndex))

                    Case Core.PropertyScope.ContentProperty
                      lobjValue = DataLookup.GetValue(Transformation.Document.Versions(lobjPropertyLookup.SourceProperty.VersionIndex).Contents(ContentElementIndex))

                    Case Else
                      lobjValue = DataLookup.GetValue(Transformation.Document)
                  End Select

                  lblnChangeMimeTypeReturn = Transformation.Document.ChangeContentMimeType(lobjValue.ToString, VersionIndex, ContentElementIndex, lstrChangeMessage)
                  If lblnChangeMimeTypeReturn = False Then
                    lpErrorMessage = lstrChangeMessage
                  End If
                  'Return New ActionResult(Me, lblnChangeMimeTypeReturn, lpErrorMessage)
                Else

                  Select Case lobjPropertyLookup.SourceProperty.PropertyScope

                    Case Core.PropertyScope.VersionProperty, Core.PropertyScope.ContentProperty

                      Dim lintVersionIndex As Integer = 0

                      For Each lobjVersion As Version In Transformation.Document.Versions

                        If lobjPropertyLookup.SourceProperty.PropertyScope = Core.PropertyScope.VersionProperty Then
                          lobjValue = DataLookup.GetValue(Transformation.Document.Versions(lintVersionIndex))

                        ElseIf lobjPropertyLookup.SourceProperty.PropertyScope = Core.PropertyScope.ContentProperty Then

                          If Transformation.Document.Versions(lintVersionIndex).Contents.Count > 0 Then
                            lobjValue = DataLookup.GetValue(Transformation.Document.Versions(lintVersionIndex).Contents(ContentElementIndex))
                            lblnChangeMimeTypeReturn = Transformation.Document.ChangeContentMimeType(lobjValue.ToString, lintVersionIndex, ContentElementIndex, lstrChangeMessage)
                            If lblnChangeMimeTypeReturn = False Then
                              lpErrorMessage = lstrChangeMessage
                            End If
                            'If lblnChangeMimeTypeReturn = False Then
                            '  ApplicationLogging.WriteLogEntry(lpErrorMessage, Reflection.MethodBase.GetCurrentMethod, TraceEventType.Warning, 62301)
                            'End If

                            lintVersionIndex += 1

                          Else
                            lpErrorMessage &= String.Format("Unable to change content mime type for element {0} of version {1}.  There is no content element present. ", ContentElementIndex, lintVersionIndex)
                          End If

                        End If

                      Next

                      'Return New ActionResult(Me, lblnChangeMimeTypeReturn, lpErrorMessage)
                    Case Else
                      lobjValue = DataLookup.GetValue(Transformation.Document)
                      lblnChangeMimeTypeReturn = Transformation.Document.ChangeContentMimeType(lobjValue.ToString, VersionIndex, ContentElementIndex, lstrChangeMessage)
                      If lblnChangeMimeTypeReturn = False Then
                        lpErrorMessage = lstrChangeMessage
                      End If
                  End Select

                End If

              Else
                lobjValue = DataLookup.GetValue(Transformation.Document)
              End If

              If lblnChangeMimeTypeReturn Is Nothing Then
                lblnChangeMimeTypeReturn = Transformation.Document.ChangeContentMimeType(lobjValue.ToString, VersionIndex, ContentElementIndex, lstrChangeMessage)
                If lblnChangeMimeTypeReturn = False Then
                  lpErrorMessage = lstrChangeMessage
                End If
              End If

              'Return New ActionResult(Me, lblnChangeMimeTypeReturn, lpErrorMessage)

              If lobjValue IsNot Nothing AndAlso lblnChangeMimeTypeReturn = False AndAlso TypeOf (DataLookup) Is DataList Then

                ' Check to see if the proposed value was listed in the ValueMap
                If DirectCast(DataLookup, DataList).MapList.OriginalExists(lobjValue.ToString) = False Then
                  lpErrorMessage = String.Format("No value map was found in the map list for file extension '{0}'.", lobjValue.ToString.ToLower)
                End If

              End If

              If lblnChangeMimeTypeReturn = True Then
                Return New ActionResult(Me, lblnChangeMimeTypeReturn, String.Empty)
              Else
                Return New ActionResult(Me, lblnChangeMimeTypeReturn, lpErrorMessage)
              End If

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
            "Specifies which content element to change."))
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