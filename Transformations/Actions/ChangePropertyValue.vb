'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Collections.ObjectModel
Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Data
Imports Documents.Data.Exceptions
Imports Documents.Exceptions
Imports Documents.Scripting
Imports Documents.Utilities

#End Region

Namespace Transformations
  ''' <summary>Action used to change the value of a property.</summary>
  <XmlInclude(GetType(ChangeContentRetrievalName)),
   XmlInclude(GetType(ChangeContentMimeType)),
   XmlInclude(GetType(ChangeDateTimePropertyValueToUTC)),
   XmlInclude(GetType(RemoveTimeFromDatePropertyValueAction)),
   XmlInclude(GetType(DecisionAction)),
   XmlInclude(GetType(AddPropertyValue))>
  <Serializable()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ChangePropertyValue
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "ChangePropertyValue"
    Friend Const PARAM_CHANGE_TYPE As String = "ChangeType"
    Friend Const PARAM_SOURCE_TYPE As String = "SourceType"
    Friend Const PARAM_ALLOW_DUPLICATES As String = "AllowDuplicates"

    Friend Const PARAM_SOURCE_PROPERTY_NAME As String = "SourcePropertyName"
    Friend Const PARAM_SOURCE_PROPERTY_SCOPE As String = "SourcePropertyScope"
    Friend Const PARAM_SOURCE_PROPERTY_TYPE As String = "SourcePropertyType"
    Friend Const PARAM_SOURCE_VERSION_INDEX As String = "SourceVersionIndex"
    Friend Const PARAM_SOURCE_VALUE_INDEX As String = "SourceValueIndex"

    Friend Const PARAM_DESTINATION_PROPERTY_NAME As String = "DestinationPropertyName"
    Friend Const PARAM_DESTINATION_PROPERTY_SCOPE As String = "DestinationPropertyScope"
    Friend Const PARAM_DESTINATION_PROPERTY_TYPE As String = "DestinationPropertyType"
    Friend Const PARAM_DESTINATION_VERSION_INDEX As String = "DestinationVersionIndex"
    Friend Const PARAM_DESTINATION_VALUE_INDEX As String = "DestinationValueIndex"

    Friend Const PARAM_MAP_LIST As String = "MapList"
    Friend Const PARAM_MAP_LIST_DELIMITER As String = "MapListDelimiter"

#End Region

#Region "Enumerations"

    ''' <summary>
    ''' Determines if the property value is to be changed via a literal value or a data
    ''' lookup.
    ''' </summary>
    Public Enum ValueSource
      ''' <summary>A literal value will be used to change the property value.</summary>
      Literal = 0
      ''' <summary>The value to be applied to the property will come from a data lookup.</summary>
      DataLookup = 1
      'ConditionalLiteral = 2
      NewGuid = 2
    End Enum

    Public Enum SourceTypeEnum
      Literal = 0
      DataParser = 1
      DataList = 2
      DataSource = 3
      NewGuid = 4
    End Enum

#End Region

#Region "Class Variables"

    Private menuValueSource As ValueSource
    Private menuSourceType As SourceTypeEnum
    Private mstrPropertyValue As Object
    Private mintVersionIndex As Integer = Transformation.TRANSFORM_ALL_VERSIONS
    'Private mstrDataMapPath As String
    Private mobjDataLookup As DataLookup
    Private mblnAllowDuplicates As Boolean
    Private mstrMapListDelimiter As String = ":"

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    <XmlAttribute()>
    Public Overrides Property Name() As String
      Get
        Try
          'Dim lobjPropertyLookup As IPropertyLookup
          If mstrName.Length < 2 Then
            Dim lstrValueSourceType As String = String.Empty
            If Me.SourceType = ValueSource.Literal Then
              If PropertyName Is Nothing OrElse PropertyValue Is Nothing Then
                Return Me.GetType.Name
              Else
                lstrValueSourceType = String.Format("Literal({0} = {1})", PropertyName, PropertyValue.ToString)
              End If
            ElseIf Me.SourceType = ValueSource.DataLookup Then
              If DataLookup Is Nothing Then
                Return Me.GetType.Name
              Else
                lstrValueSourceType = DataLookup.Name
              End If
            End If
            If lstrValueSourceType.Length > 0 Then
              mstrName = String.Format("{0}.{1}", Me.GetType.Name, lstrValueSourceType)
            Else
              mstrName = Me.GetType.Name
            End If
          End If
          Return mstrName
        Catch ex As Exception
          Return MyBase.mstrName
        End Try
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    Public Property PropertyValue() As Object
      Get
        Return mstrPropertyValue
      End Get
      Set(ByVal Value As Object)
        Try
          mstrPropertyValue = Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property VersionIndex() As Integer
      Get
        Return mintVersionIndex
      End Get
      Set(ByVal Value As Integer)
        mintVersionIndex = Value
      End Set
    End Property

    Public Property SourceType() As ValueSource
      Get
        Return menuValueSource
      End Get
      Set(ByVal Value As ValueSource)
        Try
          menuValueSource = Value
          Select Case Value
            Case ValueSource.Literal
              menuSourceType = SourceTypeEnum.Literal
          End Select
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property ChangeType As SourceTypeEnum
      Get
        Try
          Return menuSourceType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property AllowDuplicates() As Boolean
      Get
        Return mblnAllowDuplicates
      End Get
      Set(ByVal value As Boolean)
        mblnAllowDuplicates = value
      End Set
    End Property

    'Public Property DataMapPath() As String
    '  Get
    '    Return mstrDataMapPath
    '  End Get
    '  Set(ByVal Value As String)
    '    mstrDataMapPath = Value
    '  End Set
    'End Property

    Public Property DataLookup() As DataLookup
      Get
        Return mobjDataLookup
      End Get
      Set(ByVal Value As DataLookup)
        Try
          mobjDataLookup = Value
          mobjDataLookup.Action = Me
          If Value IsNot Nothing Then
            If TypeOf Value Is DataParser Then
              menuSourceType = SourceTypeEnum.DataParser
            ElseIf TypeOf Value Is DataSource Then
              menuSourceType = SourceTypeEnum.DataSource
            ElseIf TypeOf Value Is DataList Then
              menuSourceType = SourceTypeEnum.DataList
            End If
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property MapList As ObservableCollection(Of String)
      Get
        Try
          If ChangeType = SourceTypeEnum.DataList Then
            Return CType(Me.DataLookup, DataList).MapList.GetMapListItems(Me.MapListDelimiter)
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As ObservableCollection(Of String))
        Try
          If ChangeType = SourceTypeEnum.DataList AndAlso Me.DataLookup IsNot Nothing Then
            Dim lstrValueMap() As String
            For Each lstrValue As String In value
              lstrValueMap = lstrValue.Split(Me.MapListDelimiter)
              If lstrValueMap.Length = 2 Then
                CType(Me.DataLookup, DataList).MapList.Add(lstrValueMap(0), lstrValueMap(1))
              End If
            Next
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Friend Property MapListDelimiter As String
      Get
        Try
          Return mstrMapListDelimiter
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrMapListDelimiter = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
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
                   Optional ByVal lpSourceType As ValueSource = ValueSource.Literal,
                   Optional ByVal lpDataLookup As DataLookup = Nothing,
                   Optional ByVal lpLiteralPropertyValue As Object = Nothing)

      MyBase.New(lpPropertyName, ActionType.ChangePropertyValue)
      'PropertyValue = lpPropertyValue
      PropertyScope = lpPropertyScope
      VersionIndex = lpVersionIndex
      SourceType = lpSourceType

      'DataMapPath = lpDataMapPath
      If lpDataLookup IsNot Nothing Then
        DataLookup = lpDataLookup
      End If

      If lpLiteralPropertyValue IsNot Nothing Then
        PropertyValue = lpLiteralPropertyValue
      End If

    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult

      '' Change a property value
      'Dim lobjChangePropertyValueAction As ChangePropertyValue = lobjAction
      Try

        Dim lobjActionResult As ActionResult = Nothing

        If SourceExists() = False Then
          ' We were not able to verify the source property above
          If TypeOf (DataLookup) Is IPropertyLookup Then
            lpErrorMessage = String.Format("Unable to change property value of {0}, the source property {1} does not exist.",
                                           PropertyName, CType(DataLookup, IPropertyLookup).SourceProperty.PropertyName)
          Else
            lpErrorMessage = String.Format("Unable to change property value of {0}, the source does not exist.", PropertyName)
          End If
          Return New ActionResult(Me, False, lpErrorMessage)
        End If

        Select Case SourceType
          Case ChangePropertyValue.ValueSource.Literal
            Try

              ' See if this is a script or not
              If Script.IsCtScript(PropertyValue) = True Then
                PropertyValue = Script.GetValue(PropertyValue)
              End If

              If TypeOf Transformation.Target Is Document Then
                Transformation.Document.ChangePropertyValue(PropertyScope,
                                                            PropertyName,
                                                            PropertyValue,
                                                            VersionIndex)
              ElseIf TypeOf Transformation.Target Is Folder Then
                Transformation.Folder.ChangePropertyValue(PropertyName,
                                            PropertyValue)
              End If

              Return New ActionResult(Me, True)

            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              Return New ActionResult(Me, False, ex.Message)
            End Try

          Case ChangePropertyValue.ValueSource.DataLookup
            'debug.writeline(lobjChangePropertyValueAction.DataMap.SQLStatement(lpEcmDocument))
            Try

              Dim lobjValue As Object = Nothing

              If TypeOf Transformation.Target Is Document Then
                If Me.VersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
                  For lintVersionCounter As Integer = 0 To Transformation.Document.Versions.Count - 1
                    Dim lobjVersion As Version = Transformation.Document.Versions(lintVersionCounter)
                    Dim lobjProperty As ECMProperty = Nothing

                    If PropertyScope = Core.PropertyScope.DocumentProperty Then
                      lobjProperty = Transformation.Document.Properties(PropertyName)
                    Else
                      lobjProperty = lobjVersion.Properties(PropertyName)
                    End If


                    If lobjProperty Is Nothing AndAlso TypeOf (DataLookup) Is IPropertyLookup Then
                      If CType(DataLookup, IPropertyLookup).DestinationProperty.AutoCreate = True Then
                        ' We need to create the destination property
                        Dim lobjDestinationLookup As LookupProperty = CType(DataLookup, IPropertyLookup).DestinationProperty
                        Transformation.Document.CreateProperty(lobjDestinationLookup.PropertyName, , PropertyType.ecmString, lobjDestinationLookup.PropertyScope)

                        CType(DataLookup, IPropertyLookup).DestinationProperty.CreateProperty(Transformation.Document)

                        If PropertyScope = Core.PropertyScope.DocumentProperty Then
                          lobjProperty = Transformation.Document.Metadata.Item(lobjDestinationLookup.PropertyName)
                        Else
                          lobjProperty = lobjVersion.Metadata.Item(lobjDestinationLookup.PropertyName)
                        End If

                        If lobjProperty Is Nothing Then
                          ' We can't go on.
                          lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist and AutoCreate failed.",
                                     PropertyName)
                          Return New ActionResult(Me, False, lpErrorMessage)
                        End If
                      Else
                        ' We can't go on.
                        lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist and AutoCreate is set to false.",
                                   PropertyName)
                        Return New ActionResult(Me, False, lpErrorMessage)
                      End If
                    End If

                    If TypeOf (DataLookup) Is DataParser AndAlso
                      CType(DataLookup, DataParser).SourceProperty.PropertyScope = Core.PropertyScope.DocumentProperty Then
                      If lobjValue Is Nothing Then
                        ' We need to get the source value from a document property
                        If (lobjProperty.Cardinality = Cardinality.ecmSingleValued) Then
                          lobjValue = DataLookup.GetValue(Transformation.Document)
                        Else
                          lobjValue = DataLookup.GetValues(Transformation.Document)
                        End If
                      Else
                        ' We already have the source value, since it is a document property, 
                        ' it will not change with each version, just use the cached value.
                      End If
                    Else
                      ' We need to get the source value from a version property
                      If (lobjProperty.Cardinality = Cardinality.ecmSingleValued) Then
                        lobjValue = DataLookup.GetValue(lobjVersion)
                      Else
                        lobjValue = DataLookup.GetValues(lobjVersion)
                      End If
                    End If

                    Transformation.Document.ChangePropertyValue(PropertyScope, PropertyName, lobjValue, lintVersionCounter)

                  Next
                Else
                  Dim lobjProperty As ECMProperty = Nothing

                  Select Case PropertyScope
                    Case Core.PropertyScope.DocumentProperty
                      If Transformation.Document.PropertyExists(Core.PropertyScope.DocumentProperty, PropertyName) Then
                        lobjProperty = Transformation.Document.Properties(PropertyName)
                      End If
                    Case Core.PropertyScope.VersionProperty
                      If Transformation.Document.PropertyExists(Core.PropertyScope.VersionProperty, PropertyName) Then
                        lobjProperty = Transformation.Document.Versions(VersionIndex).Properties(PropertyName)
                      End If
                    Case Core.PropertyScope.ContentProperty
                      If Transformation.Document.PropertyExists(Core.PropertyScope.ContentProperty, PropertyName) Then
                        lobjProperty = Transformation.Document.GetProperty(PropertyName, Core.PropertyScope.ContentProperty, VersionIndex)
                      End If
                  End Select

                  ' Make sure the destination property exists
                  If lobjProperty Is Nothing Then
                    Dim lobjVersion As Version = Transformation.Document.Versions(VersionIndex)

                    If lobjProperty Is Nothing AndAlso TypeOf (DataLookup) Is IPropertyLookup Then
                      If CType(DataLookup, IPropertyLookup).DestinationProperty.AutoCreate = True Then
                        lobjProperty = CType(DataLookup, IPropertyLookup).DestinationProperty.CreateProperty(Transformation.Document)
                        If lobjProperty Is Nothing Then
                          ' We can't go on.
                          lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist and AutoCreate failed.",
                                     PropertyName)
                          Return New ActionResult(Me, False, lpErrorMessage)
                        End If
                      Else
                        ' We can't go on.
                        lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist and AutoCreate is set to false.",
                                   PropertyName)
                        Return New ActionResult(Me, False, lpErrorMessage)
                      End If
                      '' We were not able to get the destination property above
                      'lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist.", PropertyName)
                      'Return New ActionResult(Me, False, lpErrorMessage)
                    End If
                  End If
                  If (lobjProperty.Cardinality = Cardinality.ecmSingleValued) Then
                    lobjValue = DataLookup.GetValue(Transformation.Document)
                    Transformation.Document.ChangePropertyValue(PropertyScope, PropertyName, lobjValue, VersionIndex)
                  Else
                    lobjValue = DataLookup.GetValues(Transformation.Document)
                    If TypeOf (DataLookup) Is DataMap Then
                      Dim lobjDataMap As DataMap = CType(DataLookup, DataMap)
                      If lobjValue IsNot Nothing AndAlso TypeOf (lobjValue) Is MapList Then
                        Dim lobjMapList As MapList = CType(lobjValue, MapList)
                        If lobjDataMap.SourceColumn.ToLower.EndsWith("_mv") Then
                          ' This is a delimited value, we need to try to split it 
                          ' and add the individual values to the destination property.
                          If lobjDataMap.Delimiter IsNot Nothing AndAlso
                          lobjDataMap.Delimiter.Length > 0 Then
                            ' Split the value and add each product to the destination property
                            Dim lstrSplitValues() As String
                            For Each lobjValueMap As ValueMap In lobjMapList
                              lstrSplitValues = lobjValueMap.Replacement.Split(lobjDataMap.Delimiter)
                              For lintValueCounter As Integer = 0 To lstrSplitValues.Length - 1
                                Try
                                  Transformation.Document.AddPropertyValue(PropertyScope, PropertyName, lstrSplitValues(lintValueCounter), VersionIndex, AllowDuplicates)
                                Catch ValueExistsEx As Exceptions.ValueExistsException
                                  If lpErrorMessage.Length > 0 Then
                                    lpErrorMessage &= String.Format(", {0}", ValueExistsEx.Message)
                                  Else
                                    lpErrorMessage = ValueExistsEx.Message
                                  End If
                                End Try
                              Next
                            Next
                          Else
                            For Each lobjValueMap As ValueMap In lobjMapList
                              Try
                                Transformation.Document.AddPropertyValue(PropertyScope, PropertyName, lobjValueMap.Replacement, VersionIndex, AllowDuplicates)
                              Catch ValueExistsEx As Exceptions.ValueExistsException
                                If lpErrorMessage.Length > 0 Then
                                  lpErrorMessage &= String.Format(", {0}", ValueExistsEx.Message)
                                Else
                                  lpErrorMessage = ValueExistsEx.Message
                                End If
                              End Try
                            Next
                          End If
                        Else
                          For Each lobjValueMap As ValueMap In lobjMapList
                            Try
                              Transformation.Document.AddPropertyValue(PropertyScope, PropertyName, lobjValueMap.Replacement, VersionIndex, AllowDuplicates)
                            Catch ValueExistsEx As Exceptions.ValueExistsException
                              If lpErrorMessage.Length > 0 Then
                                lpErrorMessage &= String.Format(", {0}", ValueExistsEx.Message)
                              Else
                                lpErrorMessage = ValueExistsEx.Message
                              End If
                            End Try
                          Next
                        End If
                      Else
                        ' We can't split the value, we do not have a valid DataMap
                        lpErrorMessage = String.Format("Unable to add any values to {0}, no results were returned from the data lookup.", PropertyName)
                        Return New ActionResult(Me, False, lpErrorMessage)
                      End If

                    End If
                  End If

                End If

              ElseIf TypeOf Transformation.Target Is Folder Then

                Dim lobjProperty As ECMProperty = Nothing

                If Transformation.Folder.PropertyExists(PropertyName) Then
                  lobjProperty = Transformation.Document.Properties(PropertyName)
                End If

                ' Make sure the destination property exists
                If lobjProperty Is Nothing Then
                  Dim lobjVersion As Version = Transformation.Document.Versions(VersionIndex)

                  If lobjProperty Is Nothing AndAlso TypeOf (DataLookup) Is IPropertyLookup Then
                    If CType(DataLookup, IPropertyLookup).DestinationProperty.AutoCreate = True Then
                      lobjProperty = CType(DataLookup, IPropertyLookup).DestinationProperty.CreateProperty(Transformation.Document)
                      If lobjProperty Is Nothing Then
                        ' We can't go on.
                        lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist and AutoCreate failed.",
                                   PropertyName)
                        Return New ActionResult(Me, False, lpErrorMessage)
                      End If
                    Else
                      ' We can't go on.
                      lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist and AutoCreate is set to false.",
                                 PropertyName)
                      Return New ActionResult(Me, False, lpErrorMessage)
                    End If
                    '' We were not able to get the destination property above
                    'lpErrorMessage = String.Format("Unable to change property value of {0}, the property does not exist.", PropertyName)
                    'Return New ActionResult(Me, False, lpErrorMessage)
                  End If
                End If

                If (lobjProperty.Cardinality = Cardinality.ecmSingleValued) Then
                  lobjValue = DataLookup.GetValue(Transformation.Folder)
                  Transformation.Folder.ChangePropertyValue(PropertyName, lobjValue)
                Else
                  lobjValue = DataLookup.GetValues(Transformation.Folder)
                  If TypeOf (DataLookup) Is DataMap Then
                    Dim lobjDataMap As DataMap = CType(DataLookup, DataMap)
                    If lobjValue IsNot Nothing AndAlso TypeOf (lobjValue) Is MapList Then
                      Dim lobjMapList As MapList = CType(lobjValue, MapList)
                      If lobjDataMap.SourceColumn.ToLower.EndsWith("_mv") Then
                        ' This is a delimited value, we need to try to split it 
                        ' and add the individual values to the destination property.
                        If lobjDataMap.Delimiter IsNot Nothing AndAlso
                        lobjDataMap.Delimiter.Length > 0 Then
                          ' Split the value and add each product to the destination property
                          Dim lstrSplitValues() As String
                          For Each lobjValueMap As ValueMap In lobjMapList
                            lstrSplitValues = lobjValueMap.Replacement.Split(lobjDataMap.Delimiter)
                            For lintValueCounter As Integer = 0 To lstrSplitValues.Length - 1
                              Try
                                Transformation.Document.AddPropertyValue(PropertyScope, PropertyName, lstrSplitValues(lintValueCounter), VersionIndex, AllowDuplicates)
                              Catch ValueExistsEx As Exceptions.ValueExistsException
                                If lpErrorMessage.Length > 0 Then
                                  lpErrorMessage &= String.Format(", {0}", ValueExistsEx.Message)
                                Else
                                  lpErrorMessage = ValueExistsEx.Message
                                End If
                              End Try
                            Next
                          Next
                        Else
                          For Each lobjValueMap As ValueMap In lobjMapList
                            Try
                              Transformation.Folder.AddPropertyValue(PropertyName, lobjValueMap.Replacement, AllowDuplicates)
                            Catch ValueExistsEx As Exceptions.ValueExistsException
                              If lpErrorMessage.Length > 0 Then
                                lpErrorMessage &= String.Format(", {0}", ValueExistsEx.Message)
                              Else
                                lpErrorMessage = ValueExistsEx.Message
                              End If
                            End Try
                          Next
                        End If
                      Else
                        For Each lobjValueMap As ValueMap In lobjMapList
                          Try
                            Transformation.Folder.AddPropertyValue(PropertyName, lobjValueMap.Replacement, AllowDuplicates)
                          Catch ValueExistsEx As Exceptions.ValueExistsException
                            If lpErrorMessage.Length > 0 Then
                              lpErrorMessage &= String.Format(", {0}", ValueExistsEx.Message)
                            Else
                              lpErrorMessage = ValueExistsEx.Message
                            End If
                          End Try
                        Next
                      End If
                    Else
                      ' We can't split the value, we do not have a valid DataMap
                      lpErrorMessage = String.Format("Unable to add any values to {0}, no results were returned from the data lookup.", PropertyName)
                      Return New ActionResult(Me, False, lpErrorMessage)
                    End If

                  End If
                End If

              End If
              Return New ActionResult(Me, True)
            Catch ValueNotFoundEx As ValueNotFoundException
              lpErrorMessage = String.Format("DataMap found no value for [{0}]: {1}", PropertyName, ValueNotFoundEx.Message)
              Return New ActionResult(Me, False, lpErrorMessage)
            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              'If ex.Message.StartsWith("No value found for the expression") = True Then
              '  'Throw New Exception("DataMap found no value for [" & lobjChangePropertyValueAction.PropertyName & "]", ex)
              '  lpErrorMessage &= "DataMap found no value for [" & PropertyName & "]" & ex.Message
              '  Return New ActionResult(Me, False, lpErrorMessage)
              'Else
              'Throw New Exception("DataLookup Failed", ex)
              lpErrorMessage &= "DataLookup Failed; " & Helper.FormatCallStack(ex)
              Return New ActionResult(Me, False, lpErrorMessage)
              'End If
            Finally

            End Try

          Case ValueSource.NewGuid
            Try

              ' Assign a new guid to the property
              PropertyValue = Guid.NewGuid.ToString

              If TypeOf Transformation.Target Is Document Then
                Transformation.Document.ChangePropertyValue(PropertyScope,
                                                            PropertyName,
                                                            PropertyValue,
                                                            VersionIndex)
              ElseIf TypeOf Transformation.Target Is Folder Then
                Transformation.Folder.ChangePropertyValue(PropertyName,
                                            PropertyValue)
              End If

              Return New ActionResult(Me, True)

            Catch ex As Exception
              ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
              Return New ActionResult(Me, False, ex.Message)
            End Try
        End Select

        Return New ActionResult(Me, True)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("{0}::Execute", Me.GetType.Name))
        lpErrorMessage &= String.Format("ChangePropertyValue Failed: {0}", Helper.FormatCallStack(ex))
        Return New ActionResult(Me, False, lpErrorMessage)
      End Try

    End Function

    Public Overrides Function ToActionItem() As IActionItem
      Try
        If ChangeType = SourceTypeEnum.DataList Then
          Return New MapListActionItem(Me, CType(Me.DataLookup, DataList).MapList, Me.MapListDelimiter)
        Else
          'Return New ActionItem(Me)
          Return New MapListActionItem(Me, New MapList(), Me.MapListDelimiter)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Friend Methods"

    Friend Shared Function ValidateDateProperty(ByVal lpProperty As ECMProperty, ByVal lpDocument As Document,
                                           ByVal lpPropertyScope As Integer, ByRef lpErrorMessage As String) As Boolean

      Try

        lpErrorMessage = String.Empty

        If lpProperty.Type <> PropertyType.ecmDate Then
          ' Make sure the property is a date property
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the property is type {1}.", lpProperty.Name, lpProperty.Type.ToString)
        ElseIf lpProperty.Value Is Nothing Then
          ' Make sure the property is not null
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the current value is null.", lpProperty.Name)
        ElseIf lpProperty.Value.ToString.Length = 0 Then
          ' Make sure the property is not empty
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the current value is empty.", lpProperty.Name)
        ElseIf IsDate(lpProperty.Value) = False Then
          ' Make sure the property value is a value date value
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the current value is not a valid date.", lpProperty.Name)
        ElseIf lpDocument.PropertyExists(lpPropertyScope, lpProperty.Name, False) = False Then
          ' Make sure the property exists in the document or version
          Select Case lpPropertyScope
            Case PropertyScope.VersionProperty
              lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the property is not present in the version.", lpProperty.Name)
            Case Else
              lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the property is not present in the document.", lpProperty.Name)
          End Select
        End If

        If String.IsNullOrEmpty(lpErrorMessage) Then
          Return True
        Else
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Friend Shared Function ValidateDateProperty(ByVal lpProperty As ECMProperty, ByVal lpFolder As Folder,
                                           ByRef lpErrorMessage As String) As Boolean

      Try

        lpErrorMessage = String.Empty

        If lpProperty.Type <> PropertyType.ecmDate Then
          ' Make sure the property is a date property
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the property is type {1}.", lpProperty.Name, lpProperty.Type.ToString)
        ElseIf lpProperty.Value Is Nothing Then
          ' Make sure the property is not null
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the current value is null.", lpProperty.Name)
        ElseIf lpProperty.Value.ToString.Length = 0 Then
          ' Make sure the property is not empty
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the current value is empty.", lpProperty.Name)
        ElseIf IsDate(lpProperty.Value) = False Then
          ' Make sure the property value is a value date value
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the current value is not a valid date.", lpProperty.Name)
        ElseIf lpFolder.PropertyExists(lpProperty.Name, False) = False Then
          ' Make sure the property exists in the document or version
          lpErrorMessage = String.Format("Unable to change date value for property '{0}' to UTC, the property is not present in the folder.", lpProperty.Name)
        End If

        If String.IsNullOrEmpty(lpErrorMessage) Then
          Return True
        Else
          Return False
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Core.Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_CHANGE_TYPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_CHANGE_TYPE, Me.ChangeType, GetType(SourceTypeEnum),
            "Specifies the type of source for the property change."))
        End If

        If lobjParameters.Contains(PARAM_SOURCE_TYPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_SOURCE_TYPE, Me.SourceType, GetType(ValueSource),
            "Specifies the type of source for the property change."))
        End If

        'If ChangeType = SourceTypeEnum.Literal Then
        '  lobjParameters.AddRange(MyBase.GetDefaultParameters)
        'End If

        lobjParameters.AddRange(MyBase.GetDefaultParameters)

        Dim lobjLookupPropertyParameters As IParameters = Nothing

        If Me.DataLookup Is Nothing Then

          lobjLookupPropertyParameters = LookupProperty.GetDefaultParameters()

          For Each lobjLookupPropertyParameter As IParameter In lobjLookupPropertyParameters
            lobjLookupPropertyParameter.Name = String.Format("Source{0}", lobjLookupPropertyParameter.Name)
            lobjParameters.Add(lobjLookupPropertyParameter)
          Next

          lobjLookupPropertyParameters = LookupProperty.GetDefaultParameters()

          For Each lobjLookupPropertyParameter As IParameter In lobjLookupPropertyParameters
            lobjLookupPropertyParameter.Name = String.Format("Destination{0}", lobjLookupPropertyParameter.Name)
            lobjParameters.Add(lobjLookupPropertyParameter)
          Next

        Else
          'Select Case Me.ChangeType
          '  Case SourceTypeEnum.DataParser
          '    With CType(Me.DataLookup, DataParser)
          '      If lobjParameters.Contains(PARAM_SOURCE_PROPERTY_NAME) = False Then
          '        lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_SOURCE_PROPERTY_NAME, .Name, _
          '          "Specifies the name of the target property to modify."))
          '      End If
          '      'lobjLookupPropertyParameter.Name = String.Format("Source{0}", lobjLookupPropertyParameter.Name)
          '    End With
          'End Select
          Dim lobjLookup As IPropertyLookup = Me.DataLookup
          For Each lobjLookupPropertyParameter As IParameter In lobjLookup.GetParameters()
            lobjParameters.Add(lobjLookupPropertyParameter)
          Next
        End If

        If lobjParameters.Contains(PARAM_ALLOW_DUPLICATES) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_ALLOW_DUPLICATES, False,
            "When changing multi-valued properties, determines whether or not duplicates are allowed."))
        End If

        Dim lobjDataParserParameters As IParameters = DataParser.GetDefaultParameters
        For Each lobjDataParserParameter As IParameter In lobjDataParserParameters
          lobjParameters.Add(lobjDataParserParameter)
        Next

        If lobjParameters.Contains(PARAM_MAP_LIST) = False Then
          ' lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_MAP_LIST, PARAM_MAP_LIST, Cardinality.ecmMultiValued))
          Dim lobjMapListParameter As IParameter = ParameterFactory.Create(ParameterFactory.Create(PropertyType.ecmString, PARAM_MAP_LIST, PARAM_MAP_LIST, Cardinality.ecmMultiValued))
          '       "Contains a substitution list of property values.")
          lobjMapListParameter.Description = "Contains a substitution list of property values."

          Dim lobjMapList As ObservableCollection(Of String) = Me.MapList
          If lobjMapList IsNot Nothing Then
            lobjMapListParameter.Values.AddRange(lobjMapList)
          End If


          'Dim lobjMapListItems As IEnumerable = GetMapListItems()
          'lobjMapListParameter.Value = GetMapListItems()
          'For Each lstrValueMap As String In lobjMapListItems
          '  lobjMapListParameter.Values.Add(lstrValueMap)
          'Next

          lobjParameters.Add(lobjMapListParameter)
        End If

        If lobjParameters.Contains(PARAM_MAP_LIST_DELIMITER) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_MAP_LIST_DELIMITER, ":",
            "The delimiter to use between the original and replacement values."))
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
        ' Me.SourceType = GetEnumParameterValue(PARAM_SOURCE_TYPE, GetType(SourceTypeEnum), SourceTypeEnum.DataParser)
        Me.SourceType = GetEnumParameterValue(PARAM_SOURCE_TYPE, GetType(ValueSource), ValueSource.Literal)
        menuSourceType = GetEnumParameterValue(PARAM_CHANGE_TYPE, GetType(SourceTypeEnum), SourceTypeEnum.Literal)

        ' Figure out how to set the lookup property parameters
        Me.AllowDuplicates = GetBooleanParameterValue(PARAM_ALLOW_DUPLICATES, False)

        Select Case Me.ChangeType
          Case SourceTypeEnum.Literal

          Case SourceTypeEnum.DataParser
            InitializeDataParserFromParameters()

          Case SourceTypeEnum.DataList
            InitializeDataListFromParameters()

          Case SourceTypeEnum.DataSource
            Me.DataLookup = New DataSource

        End Select

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub InitializeDataParserFromParameters()
      Try
        Dim lobjDataParser As New DataParser
        'Dim lstrPropertyName As String
        'Dim lenuPropertyScope As PropertyScope
        'Dim lenuPropertyType As PropertyType
        'Dim lenuVersionIndex As Long
        'Dim lenuValueIndex As PropertyType

        '' Build the source property
        'lstrPropertyName = GetStringParameterValue(PARAM_SOURCE_PROPERTY_NAME, Me.PropertyName)
        'lenuPropertyScope = GetEnumParameterValue(PARAM_SOURCE_PROPERTY_SCOPE, GetType(PropertyScope), PropertyScope.VersionProperty)
        'lenuPropertyType = GetEnumParameterValue(PARAM_SOURCE_PROPERTY_TYPE, GetType(PropertyType), PropertyType.ecmString)
        'lenuVersionIndex = GetParameterValue(PARAM_SOURCE_VERSION_INDEX, 0)
        'lenuValueIndex = GetEnumParameterValue(PARAM_SOURCE_VALUE_INDEX, GetType(Values.ValueIndexEnum), Values.ValueIndexEnum.First)

        'Dim lobjSourceLookupProperty As New LookupProperty(lstrPropertyName, lenuPropertyScope, lenuVersionIndex)
        'lobjSourceLookupProperty.Type = lenuPropertyType
        'lobjSourceLookupProperty.ValueIndex = lenuValueIndex

        Dim lobjSourceLookupProperty As LookupProperty = GetLookupPropertyFromParameters(RepositoryScope.Source)

        ' Build the destination property
        'lstrPropertyName = GetStringParameterValue(PARAM_DESTINATION_PROPERTY_NAME, Me.PropertyName)
        'lenuPropertyScope = GetEnumParameterValue(PARAM_DESTINATION_PROPERTY_SCOPE, GetType(PropertyScope), PropertyScope.VersionProperty)
        'lenuPropertyType = GetEnumParameterValue(PARAM_DESTINATION_PROPERTY_TYPE, GetType(PropertyType), PropertyType.ecmString)
        'lenuVersionIndex = GetParameterValue(PARAM_DESTINATION_VERSION_INDEX, 0)
        'lenuValueIndex = GetEnumParameterValue(PARAM_DESTINATION_VALUE_INDEX, GetType(Values.ValueIndexEnum), Values.ValueIndexEnum.First)

        'Dim lobjDestinationLookupProperty As New LookupProperty(lstrPropertyName, lenuPropertyScope, lenuVersionIndex)
        'lobjDestinationLookupProperty.Type = lenuPropertyType
        'lobjDestinationLookupProperty.ValueIndex = lenuValueIndex

        Dim lobjDestinationLookupProperty As LookupProperty = GetLookupPropertyFromParameters(RepositoryScope.Destination)

        lobjDataParser.SourceProperty = lobjSourceLookupProperty
        lobjDataParser.DestinationProperty = lobjDestinationLookupProperty

        Me.VersionIndex = lobjDestinationLookupProperty.VersionIndex

        Dim lobjPartBuilder As New StringBuilder

        Dim lenuParseStyle As DataParser.PartEnum = GetEnumParameterValue(DataParser.PARAM_PARSE_STYLE, GetType(DataParser.PartEnum), DataParser.PartEnum.Complete)
        Dim lstrParseArgs As String = GetStringParameterValue(DataParser.PARAM_PARSE_ARGUMENTS, String.Empty)

        lobjPartBuilder.Append(lenuParseStyle.ToString())

        If Not String.IsNullOrEmpty(lstrParseArgs) Then
          lobjPartBuilder.AppendFormat(":{0}", lstrParseArgs)
        End If

        lobjDataParser.Part = lobjPartBuilder.ToString()

        Me.DataLookup = lobjDataParser

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Sub InitializeDataListFromParameters()
      Try
        Me.DataLookup = New DataList
        Dim lobjSourceLookupProperty As LookupProperty = GetLookupPropertyFromParameters(RepositoryScope.Source)
        Dim lobjDestinationLookupProperty As LookupProperty = GetLookupPropertyFromParameters(RepositoryScope.Destination)
        Dim lobjMapList As New MapList



      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function GetLookupPropertyFromParameters(lpScope As RepositoryScope) As LookupProperty
      Try
        Dim lobjDataParser As New DataParser
        Dim lstrPropertyName As String
        Dim lenuPropertyScope As PropertyScope
        Dim lenuPropertyType As PropertyType
        Dim lenuVersionIndex As Long
        Dim lenuValueIndex As PropertyType

        Dim lobjLookupProperty As LookupProperty

        Select Case lpScope
          Case RepositoryScope.Source
            ' Build the source property
            lstrPropertyName = GetStringParameterValue(PARAM_SOURCE_PROPERTY_NAME, Me.PropertyName)
            lenuPropertyScope = GetEnumParameterValue(PARAM_SOURCE_PROPERTY_SCOPE, GetType(PropertyScope), PropertyScope.VersionProperty)
            lenuPropertyType = GetEnumParameterValue(PARAM_SOURCE_PROPERTY_TYPE, GetType(PropertyType), PropertyType.ecmString)
            lenuVersionIndex = GetParameterValue(PARAM_SOURCE_VERSION_INDEX, 0)
            lenuValueIndex = GetEnumParameterValue(PARAM_SOURCE_VALUE_INDEX, GetType(Values.ValueIndexEnum), Values.ValueIndexEnum.First)

          Case RepositoryScope.Destination
            ' Build the destination property
            lstrPropertyName = GetStringParameterValue(PARAM_DESTINATION_PROPERTY_NAME, Me.PropertyName)
            lenuPropertyScope = GetEnumParameterValue(PARAM_DESTINATION_PROPERTY_SCOPE, GetType(PropertyScope), PropertyScope.VersionProperty)
            lenuPropertyType = GetEnumParameterValue(PARAM_DESTINATION_PROPERTY_TYPE, GetType(PropertyType), PropertyType.ecmString)
            lenuVersionIndex = GetParameterValue(PARAM_DESTINATION_VERSION_INDEX, 0)
            lenuValueIndex = GetEnumParameterValue(PARAM_DESTINATION_VALUE_INDEX, GetType(Values.ValueIndexEnum), Values.ValueIndexEnum.First)

          Case Else
            Throw New ArgumentOutOfRangeException(NameOf(lpScope))
        End Select

        lobjLookupProperty = New LookupProperty(lstrPropertyName, lenuPropertyScope, lenuVersionIndex) With {
          .Type = lenuPropertyType,
          .ValueIndex = lenuValueIndex
        }

        Return lobjLookupProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Friend Function GetTargetProperty() As IProperty
      Try
        If TypeOf Transformation.Target Is Document Then
          Return Transformation.Document.GetProperty(Me.PropertyName, Me.PropertyScope, Me.VersionIndex)
        ElseIf TypeOf Transformation.Target Is Folder Then
          Return Transformation.Folder.GetProperty(Me.PropertyName)
        Else
          Throw New InvalidTransformationTargetException
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    'Protected Function GetMapListItems() As ObservableCollection(Of String)
    '  Try
    '    Dim lobjStringList As New ObservableCollection(Of String)

    '    If ChangeType = SourceTypeEnum.DataList Then
    '      Dim lobjMapList As MapList = CType(Me.DataLookup, DataList).MapList
    '      'For Each lobjValueMap As ValueMap In lobjMapList
    '      '  lobjStringList.Add(String.Format("{0}{1}{2}", lobjValueMap.Original, Me.MapListDelimiter, lobjValueMap.Replacement))
    '      'Next
    '    End If

    '    Return lobjStringList

    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Function

    Protected Friend Function SourceExists() As Boolean
      Try

        If TypeOf Transformation.Target Is Document Then
          Select Case Me.SourceType
            Case ValueSource.DataLookup
              If TypeOf (DataLookup) Is IPropertyLookup Then
                If CType(DataLookup, IPropertyLookup).SourceProperty IsNot Nothing Then
                  Dim lobjSourcePropertyScope As PropertyScope =
                    CType(DataLookup, IPropertyLookup).SourceProperty.PropertyScope
                  Return DataLookup.SourceExists(GetMetaHolder(lobjSourcePropertyScope))
                Else
                  Return DataLookup.SourceExists(GetMetaHolder())
                End If
              Else
                Return DataLookup.SourceExists(GetMetaHolder())
              End If
            Case ValueSource.Literal
              Return True
            Case Else
              Return True
          End Select
        ElseIf TypeOf Transformation.Target Is Folder Then
          Select Case Me.SourceType
            Case ValueSource.DataLookup
              If TypeOf (DataLookup) Is IPropertyLookup Then
                If CType(DataLookup, IPropertyLookup).SourceProperty IsNot Nothing Then
                  'Dim lobjSourcePropertyScope As PropertyScope = _
                  '  CType(DataLookup, IPropertyLookup).SourceProperty.PropertyScope
                  'Return DataLookup.SourceExists(GetMetaHolder(lobjSourcePropertyScope))
                  Return DataLookup.SourceExists(GetFolderMetaHolder())
                Else
                  Return DataLookup.SourceExists(GetFolderMetaHolder())
                End If
              Else
                Return DataLookup.SourceExists(GetFolderMetaHolder())
              End If
            Case ValueSource.Literal
              Return True
            Case Else
              Return True
          End Select
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return False
      End Try
    End Function

    Protected Overrides Function GetMetaHolder(ByVal lpPropertyScope As PropertyScope) As IMetaHolder
      Try
        Dim lobjMetaHolder As IMetaHolder = Nothing

        Select Case lpPropertyScope
          Case Core.PropertyScope.DocumentProperty
            lobjMetaHolder = Transformation.Document
          Case Core.PropertyScope.VersionProperty
            If VersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
              lobjMetaHolder = Transformation.Document.GetFirstVersion
            Else
              lobjMetaHolder = Transformation.Document.Versions(VersionIndex)
            End If
          Case Core.PropertyScope.ContentProperty
            If VersionIndex = Transformation.TRANSFORM_ALL_VERSIONS Then
              lobjMetaHolder = Transformation.Document.GetFirstVersion.Contents(0)
            Else
              If Transformation.Document.Versions(VersionIndex).Contents.Count > 0 Then
                lobjMetaHolder = Transformation.Document.Versions(VersionIndex).Contents(0)
              End If
            End If
          Case Else
            lobjMetaHolder = Transformation.Document
        End Select

        Return lobjMetaHolder

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller.
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Shadows Function DebuggerIdentifier() As String
      Try
        'Dim lobjStringBuilder As New StringBuilder
        'Dim lstrSourceType As String = String.Empty

        'Select Case SourceType
        '  Case ValueSource.Literal
        '    lstrSourceType = SourceType.ToString
        '  Case ValueSource.DataLookup
        '    If DataLookup IsNot Nothing Then
        '      lstrSourceType = DataLookup.GetType.Name
        '    Else
        '      lstrSourceType = SourceType.ToString
        '    End If
        'End Select

        'If Name Is Nothing OrElse Name.Length = 0 Then
        '  lobjStringBuilder.AppendFormat("{0}.{1}", Me.GetType.Name, lstrSourceType)
        'Else
        '  lobjStringBuilder.AppendFormat("{0}.{1}: {2}", Me.GetType.Name, lstrSourceType, Name)
        'End If

        'Return lobjStringBuilder.ToString

        Return Name

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace