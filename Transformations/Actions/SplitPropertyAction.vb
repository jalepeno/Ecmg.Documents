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
  ''' Split a property value based upon a delimiter (i.e. FolderPath)
  ''' </summary>
  ''' <remarks>This action is only supported for string properties</remarks>
  <Serializable()>
  Public Class SplitPropertyAction
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "SplitProperty"
    Private Const PARAM_DELIMITER As String = "Delimiter"
    Private Const PARAM_START_POSITION As String = "StartPosition"
    Private Const PARAM_TRIM_PREFIX As String = "TrimPrefix"

#End Region

#Region "Class Variables"

    Private mstrDelimiter As String = String.Empty
    Private mstrTrimPrefix As String = String.Empty
    Private mintStartPosition As Integer = 1
    Private mobjSourceProperty As LookupProperty
    Private mobjDestinationProperties As New LookupProperties
    'Private menuValueIndex As Values.ValueIndexEnum = Values.ValueIndexEnum.First

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    ''' <summary>
    ''' The delimiter used to split the property value.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Delimiter() As String
      Get
        Return mstrDelimiter
      End Get
      Set(ByVal value As String)
        Try
          mstrDelimiter = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the value to trim from the front 
    ''' of the source property before the split operation
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TrimPrefix() As String
      Get
        Return mstrTrimPrefix
      End Get
      Set(ByVal value As String)
        mstrTrimPrefix = value
      End Set
    End Property

    ''' <summary>
    ''' The position from which to start the split operation.
    ''' </summary>
    ''' <value>The position from which to start.  To start at the beginning, specify 1.</value>
    ''' <returns></returns>
    ''' <remarks>
    ''' If there is a TrimPrefix specified, the trim operation will run first.  
    ''' In that case the StartPosition will start after the end of the TrimPrefix.
    ''' </remarks>
    Public Property StartPosition() As Integer
      Get
        Return mintStartPosition
      End Get
      Set(ByVal value As Integer)
        mintStartPosition = value
      End Set
    End Property

    '''' <summary>
    '''' Gets or sets the desired value index.  In the case of a multi-valued source property, 
    '''' we will only attempt to split one of the values.  
    '''' The ValueIndex determines which one we will split.
    '''' </summary>
    '''' <value></value>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Public Property ValueIndex() As Values.ValueIndexEnum
    '  Get
    '    Return menuValueIndex
    '  End Get
    '  Set(ByVal value As Values.ValueIndexEnum)
    '    menuValueIndex = value
    '  End Set
    'End Property

    ''' <summary>The property containing the source value.</summary>
    Public Property SourceProperty() As LookupProperty
      Get
        Return mobjSourceProperty
      End Get
      Set(ByVal Value As LookupProperty)
        mobjSourceProperty = Value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the collection of destination 
    ''' properties to insert the split values into.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DestinationProperties() As LookupProperties
      Get
        Return mobjDestinationProperties
      End Get
      Set(ByVal value As LookupProperties)
        mobjDestinationProperties = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(ActionType.ChangePropertyValue)
    End Sub

#End Region

#Region "Public Overrides"

    Public Shadows Property PropertyName() As String
      Get
        If SourceProperty IsNot Nothing Then
          Return SourceProperty.PropertyName
        Else
          Return MyBase.PropertyName
        End If
      End Get
      Set(ByVal value As String)
        If SourceProperty IsNot Nothing Then
          SourceProperty.PropertyName = value
        Else
          MyBase.PropertyName = value
        End If
      End Set
    End Property

    Public Shadows Property PropertyScope() As Core.PropertyScope
      Get
        If SourceProperty IsNot Nothing Then
          Return SourceProperty.PropertyScope
        Else
          Return MyBase.PropertyScope
        End If
      End Get
      Set(ByVal value As Core.PropertyScope)
        If SourceProperty IsNot Nothing Then
          SourceProperty.PropertyScope = value
        Else
          MyBase.PropertyScope = value
        End If
      End Set
    End Property
#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult

      Dim lstrErrorMessage As String = String.Empty

      Try

        ' 1. Trim off the specified TrimPrefix from the front of the source property
        ' 2. Remove any characters in front of the start position
        ' 3. Split the source property using the specified delimiter 
        ' 4. Then output the separate values into the destination properties

        Dim lobjSourceProperty As ECMProperty = Nothing
        Dim lstrOriginalSourceValue As String = String.Empty
        Dim lstrPreparedSourceValue As String = String.Empty
        Dim lstrSplitValues As String()
        Dim lobjDestinationLookup As LookupProperty = Nothing
        Dim lobjDestinationProperty As ECMProperty = Nothing
        Dim lblnCreatePropertySuccess As Boolean = False

        ' Start with a clean error message
        lpErrorMessage = String.Empty

        If SourceExists() = False Then
          ' We were not able to verify the source property above
          If SourceProperty IsNot Nothing Then
            lstrErrorMessage = String.Format("Unable to split property value of {0}, the source does not exist.", SourceProperty.Name)
          Else
            lstrErrorMessage = "Unable to split property value, the source does not exist."
          End If
          lpErrorMessage &= String.Format("; {0}", lstrErrorMessage)
          Return New ActionResult(Me, False, lstrErrorMessage)
        End If

        ' Make sure we have at least one destination property
        If DestinationProperties.Count = 0 Then
          lstrErrorMessage = "Unable to execute SplitPropertyAction, at least one destination property must be specified."
          lpErrorMessage &= lstrErrorMessage
          Return New ActionResult(Me, False, lstrErrorMessage)
        End If

        ' Get the source property
        Try
          lobjSourceProperty = DataLookup.GetProperty(SourceProperty, Me)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          Return New ActionResult(Me, False, ex.Message)
        End Try

        ' Make sure we were able to get the source property
        If lobjSourceProperty Is Nothing Then
          lstrErrorMessage = String.Format("Unable to execute SplitPropertyAction, the property '{0}' could not be found.", SourceProperty.Name)
          lpErrorMessage &= lstrErrorMessage
          Return New ActionResult(Me, False, lstrErrorMessage)
        End If

        ' Make sure we are dealing with a string property
        If lobjSourceProperty.Type <> PropertyType.ecmString Then
          Throw New Exceptions.InvalidPropertyTypeException(PropertyType.ecmString, lobjSourceProperty)
        End If

        ' Get the source property value
        If lobjSourceProperty.Cardinality = Cardinality.ecmSingleValued Then
          lstrOriginalSourceValue = lobjSourceProperty.Value
        Else ' Get the value from one of the multiple values
          ' Make sure we have at least one value
          If lobjSourceProperty.Values.Count = 0 Then
            lstrErrorMessage = String.Format("Unable to execute SplitPropertyAction, the multi-valued property '{0}' has no values.", SourceProperty.Name)
            lpErrorMessage &= lstrErrorMessage
            Return New ActionResult(Me, False, lstrErrorMessage)
          End If

          Select Case SourceProperty.ValueIndex
            Case Values.ValueIndexEnum.First
              ' Get the first value
              lstrOriginalSourceValue = lobjSourceProperty.Values.GetFirstValue
            Case Values.ValueIndexEnum.Last
              ' Get the last value
              lstrOriginalSourceValue = lobjSourceProperty.Values.GetLastValue
            Case Else
              ' Get the specific value requested by index
              lstrOriginalSourceValue = lobjSourceProperty.Values(SourceProperty.ValueIndex)
          End Select
        End If

        ' Remove the prefix, if any
        lstrPreparedSourceValue = lstrOriginalSourceValue.TrimStart(TrimPrefix)

        ' Remove any characters that come before the start position
        lstrPreparedSourceValue = lstrPreparedSourceValue.Remove(0, StartPosition - 1)

        ' Split the values
        If Delimiter.ToLower.Contains("space") Then
          lstrSplitValues = lstrPreparedSourceValue.Split(" ")
        Else
          lstrSplitValues = lstrPreparedSourceValue.Split(Delimiter)
        End If

        If TypeOf Transformation.Target Is Document Then

          Dim lobjCurrentDocument As Document = Me.Transformation.Document
          ' Assign the split values to the specified destination properties
          For lintValueCounter As Integer = 0 To lstrSplitValues.Length - 1
            If DestinationProperties.Count > lintValueCounter Then

              ' Clear lobjDestinationProperty
              lobjDestinationProperty = Nothing

              ' Try to get the destination property
              lobjDestinationLookup = DestinationProperties(lintValueCounter)

              Try
                lobjDestinationProperty = Transformations.DataLookup.TryGetProperty(lobjDestinationLookup, Me)
              Catch ex As Exception
                ' The property could not be found, We will try to create it if necessary below
              End Try

              If lobjDestinationProperty Is Nothing AndAlso lobjDestinationLookup.AutoCreate = True Then
                ' We will create the property
                lobjDestinationProperty = lobjDestinationLookup.CreateProperty(lobjCurrentDocument)
                lobjDestinationProperty.SetPersistence(lobjDestinationLookup.Persistent)
              End If
            End If

            If lobjDestinationProperty IsNot Nothing Then
              If lobjDestinationLookup.VersionIndex = Document.ALL_VERSIONS Then
                ' We need to loop through each version
                For Each lobjVersion As Version In lobjCurrentDocument.Versions
                  lobjDestinationProperty = lobjVersion.Properties(lobjDestinationLookup.PropertyName)
                  lobjDestinationProperty.ChangePropertyValue(lstrSplitValues(lintValueCounter))
                Next
              Else
                ' We need to get the specified version
                lobjDestinationProperty = lobjCurrentDocument.GetProperty(lobjDestinationLookup.PropertyName, lobjDestinationLookup.VersionIndex)
                lobjDestinationProperty.ChangePropertyValue(lstrSplitValues(lintValueCounter))
              End If

            Else
              ' We were unable to get or create the destination
              lstrErrorMessage = String.Format("Unable to add value '{1}' to '{0}', the destination property '{0}' does not exist and AutoCreate is set to false.",
                                               lobjDestinationLookup.PropertyName, lstrSplitValues(lintValueCounter))
              lpErrorMessage &= lstrErrorMessage
              Return New ActionResult(Me, False, lstrErrorMessage)
            End If
          Next

          Return New ActionResult(Me, True)
        ElseIf TypeOf Transformation.Target Is Folder Then

          Dim lobjCurrentFolder As Folder = Me.Transformation.Folder

          ' Assign the split values to the specified destination properties
          For lintValueCounter As Integer = 0 To lstrSplitValues.Length - 1
            If DestinationProperties.Count > lintValueCounter Then

              ' Clear lobjDestinationProperty
              lobjDestinationProperty = Nothing

              ' Try to get the destination property
              lobjDestinationLookup = DestinationProperties(lintValueCounter)

              Try
                lobjDestinationProperty = Transformations.DataLookup.TryGetProperty(lobjDestinationLookup, Me)
              Catch ex As Exception
                ' The property could not be found, We will try to create it if necessary below
              End Try

              If lobjDestinationProperty Is Nothing AndAlso lobjDestinationLookup.AutoCreate = True Then
                ' We will create the property
                lobjDestinationProperty = lobjDestinationLookup.CreateProperty(lobjCurrentFolder)
                lobjDestinationProperty.SetPersistence(lobjDestinationLookup.Persistent)
              End If
            End If

            If lobjDestinationProperty IsNot Nothing Then
              lobjDestinationProperty = lobjCurrentFolder.GetProperty(lobjDestinationLookup.PropertyName)
              lobjDestinationProperty.ChangePropertyValue(lstrSplitValues(lintValueCounter))

            Else
              ' We were unable to get or create the destination
              lstrErrorMessage = String.Format("Unable to add value '{1}' to '{0}', the destination property '{0}' does not exist and AutoCreate is set to false.",
                                               lobjDestinationLookup.PropertyName, lstrSplitValues(lintValueCounter))
              lpErrorMessage &= lstrErrorMessage
              Return New ActionResult(Me, False, lstrErrorMessage)
            End If
          Next

          Return New ActionResult(Me, True)
        Else
          Throw New InvalidTransformationTargetException
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Execute", Me.GetType.Name))
        lstrErrorMessage = String.Format("SplitPropertyAction Failed: {0}", Helper.FormatCallStack(ex))
        lpErrorMessage &= lstrErrorMessage
        Return New ActionResult(Me, False, lstrErrorMessage)
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Friend Function SourceExists() As Boolean
      Try

        Return SourceExists(GetMetaHolder())

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try
    End Function

    Protected Friend Function SourceExists(ByVal lpMetaHolder As IMetaHolder) As Boolean
      Try

        If Me.SourceProperty IsNot Nothing AndAlso lpMetaHolder.Metadata.PropertyExists(SourceProperty.PropertyName, False) Then
          Return True
        Else
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try
    End Function

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_DELIMITER) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_DELIMITER, String.Empty,
            "The delimiter used to split the property value."))
        End If

        If lobjParameters.Contains(PARAM_START_POSITION) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_START_POSITION, 1,
            "The position from which to start the split operation."))
        End If

        If lobjParameters.Contains(PARAM_TRIM_PREFIX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_TRIM_PREFIX, String.Empty,
            "The value to trim from the front of the source property before the split operation."))
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
        Me.Delimiter = GetStringParameterValue(PARAM_DELIMITER, String.Empty)
        Me.StartPosition = GetParameterValue(PARAM_START_POSITION, 1)
        Me.TrimPrefix = GetStringParameterValue(PARAM_TRIM_PREFIX, String.Empty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"


    '''' <summary>
    '''' Gets the actual property
    '''' </summary>
    '''' <returns></returns>
    '''' <remarks></remarks>
    'Private Function GetSourceECMProperty() As ECMProperty
    '  Try

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

#End Region

  End Class

End Namespace
