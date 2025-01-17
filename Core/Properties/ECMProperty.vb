'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Globalization
Imports System.Text
Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Serialization
Imports Documents.TypeConverters
Imports Documents.Utilities
Imports Newtonsoft.Json

#End Region

Namespace Core

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ECMProperty
    Implements IProperty
    Implements IJsonSerializable(Of ECMProperty)
    Implements IXmlSerializable
    Implements IDisposable
    Implements INotifyPropertyChanged

#Region "Class Variables"

    Private menuType As PropertyType = PropertyType.ecmString
    Private menuCardinalityType As Cardinality = Cardinality.ecmSingleValued
    Private mstrName As String = String.Empty
    Private mstrPackedName As String = String.Empty
    Private mstrSystemName As String = String.Empty
    Private mobjValue As Object 'New Value
    Private mobjValues As New Values(Me)
    Private mstrID As String = String.Empty
    Private mobjDefaultValue As Object 'New Value
    Private mblnPersistent As Boolean = True
    Private mstrDescription As String = String.Empty
    Private mstrDisplayName As String = String.Empty
    Private mstrXsiType As String = String.Empty
    Protected mstrEnumTypeName As String = String.Empty
    Protected mobjEnumType As Type

    Protected mobjStandardValues As IEnumerable


#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Defines the data type of the value to be stored in an ECMProperty object.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Category("Properties"), DefaultValue(PropertyType.ecmString),
    Description("Defines the data type of the value to be stored in an ECMProperty object."),
    DisplayName("Type"), TypeConverter(GetType(EnumDescriptionTypeConverter))>
    Public Property Type() As PropertyType Implements IProperty.Type
      Get
        Return menuType
      End Get
      Set(ByVal value As PropertyType)
        Try
          menuType = value

          If value = PropertyType.ecmBoolean Then
            mobjStandardValues = New List(Of String)
            With CType(mobjStandardValues, List(Of String))
              .Add("True")
              .Add("False")
            End With

          End If

          OnPropertyChanged("Type")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Set
    End Property


    ''' <summary>
    ''' Defines the cardinality of an ECMProperty object, 
    ''' this is expressed as either single-valued or multi-valued.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>A single-valued property may only store one valued, 
    ''' whereas a multi-valued property may store multiple values for a single object instance.</remarks>
    <Category("Properties"), DefaultValue(Cardinality.ecmSingleValued),
    Description("An enumerated property describing whether or not the property is single or multi valued."),
    DisplayName("Cardinality"), TypeConverter(GetType(EnumDescriptionTypeConverter))>
    Public Property Cardinality() As Cardinality Implements IProperty.Cardinality
      Get
        Return menuCardinalityType
      End Get
      Set(ByVal value As Cardinality)
        Try
          menuCardinalityType = value
          OnPropertyChanged("Cardinality")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Set
    End Property

    ''' <summary>
    ''' The name of the property.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Usually associated with the display name for the property.</remarks>
    Public Property Name() As String Implements IProperty.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        Try
          mstrName = value
          If String.IsNullOrEmpty(ID) Then
            ID = value
          End If
          OnPropertyChanged("Name")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Returns the property name without any spaces
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PackedName() As String
      Get
        mstrPackedName = GetPackedName(mstrName)
        Return mstrPackedName
      End Get
    End Property

    ''' <summary>
    ''' The system name of the property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>In many cases there is a system name for 
    ''' properties that is different than the display name</remarks>
    Public Property SystemName As String Implements IProperty.SystemName
      Get
        Return mstrSystemName
      End Get
      Set(ByVal value As String)
        Try
          mstrSystemName = value
          OnPropertyChanged("SystemName")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Overridable ReadOnly Property HasStandardValues As Boolean Implements IProperty.HasStandardValues
      Get
        Try
          If mobjStandardValues Is Nothing Then
            Return False
          Else
            Return True
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property


    Public Overridable ReadOnly Property StandardValues As IEnumerable Implements IProperty.StandardValues
      Get
        Try
          Return mobjStandardValues
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property


    Public ReadOnly Property XsiType As String
      Get
        Return mstrXsiType
      End Get
    End Property

    ''' <summary>
    ''' The actual value assigned to the property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If the Cardinality is single-valued, the property value is stored here in the Value.  If the Cardinality is multi-valued then each value is stored in the Values collection.</remarks>
    Public Property Value() As Object Implements IProperty.Value
      Get
        Try
          'If mobjValue.Value IsNot Nothing Then
          If mobjValue IsNot Nothing Then
            'Select Case mobjValue.Value.GetType.Name
            Select Case mobjValue.GetType.Name
              Case "Value"
                Return mobjValue.ToString(mobjValue.Value)
#If Not SILVERLIGHT = 1 Then
              Case "XmlNode[]"
                ' Get the first element
                'Return CType(mobjValue.Value(0), Xml.XmlText).Value
                Return CType(mobjValue(0), Object).Value
#End If
              Case "String"
                Select Case Type
                  Case PropertyType.ecmDate
                    ' If it is supposed to be a date then coerce it to a date.
                    If IsDate(mobjValue) Then
                      Return Date.Parse(mobjValue)
                    Else
                      Return mobjValue
                    End If
                  Case Else
                    Return mobjValue
                End Select
              Case Else
                'Return mobjValue.Value
                Return mobjValue
            End Select
            'If mobjValue.Value.GetType.Name = "Value" Then
            '  Return Core.Value.ToString(mobjValue.Value)
            'Else
            '  Return mobjValue.Value
            'End If
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Object)
        Try
          If Cardinality <> Core.Cardinality.ecmSingleValued AndAlso
                 Helper.IsDeserializationBasedCall = False AndAlso
                   ((Helper.CallStackContainsMethodName("ChangePropertyCardinality") = False) _
                   AndAlso Helper.CallStackContainsMethodName("CreateProperty") = False) Then

            Throw New InvalidOperationException(
            String.Format("Unable to set the Value of property {0}, the Cardinality is {1}, please call Values.Add(value) instead.",
                          Name, Cardinality.ToString))
          End If
          If TypeOf (value) Is Value Then
            mobjValue = value.Value
          ElseIf TypeOf value Is String Then
            Select Case Type
              Case PropertyType.ecmDate
                ' If it is supposed to be a date then coerce it to a date.
                If IsDate(value) Then
                  mobjValue = Date.Parse(value)
                End If
              Case Else
                mobjValue = value
            End Select
          Else
            mobjValue = value
          End If
          OnPropertyChanged("Value")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property ValueString As String
      Get
        Try
          If Me.Value IsNot Nothing Then
            Return Me.Value.ToString
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          If Not String.IsNullOrEmpty(value) Then
            Me.Value = value
          Else
            Me.Value = Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The collection of values assigned to the property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If the Cardinality is multi-valued then each value is stored in here the Values collection.  If the Cardinality is single-valued, the property value is stored in the Value.</remarks>
    Public Property Values() As Values
      Get
        Return mobjValues
      End Get
      Set(ByVal value As Values)
        Try
          If Cardinality <> Core.Cardinality.ecmMultiValued AndAlso
                    Helper.CallStackContainsMethodName("ChangePropertyCardinality") = False Then

            Throw New InvalidOperationException(
            String.Format("Unable to set the Value of property {0}, the Cardinality is {1}, please use the Value property instead.",
                          Name, Cardinality.ToString))
          End If
          mobjValues = value
          mobjValues.SetParentProperty(Me)
          OnPropertyChanged("Values")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property ValuesEx As Object Implements IProperty.Values
      Get
        Return mobjValues
      End Get
      Set(ByVal value As Object)
        Try
          mobjValues = value
          OnPropertyChanged("Values")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' The ID of the property.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Usually associated with the system identifier for the property.</remarks>
    <XmlAttribute()>
    Public Property ID() As String
      Get
        Return mstrID
      End Get
      Set(ByVal value As String)
        Try
          mstrID = value
          OnPropertyChanged("ID")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not the property currently has a value assigned
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property HasValue() As Boolean Implements IProperty.HasValue
      Get
        Try
          Return PropertyHasValue()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, Me.Name)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' The default value to be assigned to the property for a given object instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    Public Property DefaultValue() As Object Implements IProperty.DefaultValue
      Get
        If mobjDefaultValue IsNot Nothing Then
          Return mobjDefaultValue
        Else
          Return String.Empty
        End If
      End Get
      Set(ByVal value As Object)
        Try
          mobjDefaultValue = value
          OnPropertyChanged("DefaultValue")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets a value specifiying whether or not the property 
    ''' should persist beyond the scope of the transaction 
    ''' that created it.  It is used to support the 
    ''' definition of temporary properties.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Persistent() As Boolean Implements IProperty.Persistent
      Get
        Return mblnPersistent
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets a description for this property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property Description() As String Implements IProperty.Description
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        Try
          mstrDescription = value
          OnPropertyChanged("Description")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    <XmlAttribute()>
    Public Property DisplayName() As String Implements IProperty.DisplayName
      Get
        If String.IsNullOrEmpty(mstrDisplayName) Then
          mstrDisplayName = Helper.CreateDisplayName(Me.Name)
        End If
        Return mstrDisplayName
      End Get
      Set(ByVal value As String)
        Try
          mstrDisplayName = value
          OnPropertyChanged("DisplayName")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Protected Properties"

    Protected Friend Property EnumTypeName As String
      Get
        Try
          If String.IsNullOrEmpty(mstrEnumTypeName) Then
            If Me.EnumType IsNot Nothing Then
              mstrEnumTypeName = EnumType.Name
            End If
          End If
          Return mstrEnumTypeName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          mstrEnumTypeName = value
          If mobjEnumType Is Nothing Then
            mobjEnumType = Helper.GetTypeFromAssembly(Reflection.Assembly.GetExecutingAssembly, mstrEnumTypeName)
            If mobjEnumType IsNot Nothing Then
              Dim lobjEnumDictionary As IDictionary(Of String, Integer) = Helper.EnumerationDictionary(mobjEnumType)
              mobjStandardValues = lobjEnumDictionary.Keys
              mstrEnumTypeName = EnumType.Name
            End If
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Protected Friend Property EnumType As Type
      Get
        Try
          Return mobjEnumType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Type)
        Try
          mobjEnumType = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    ''' <summary>
    ''' Intended for internal use only by the framework
    ''' </summary>
    ''' <param name="lpInternalUseOnly"></param>
    ''' <remarks></remarks>
    Protected Friend Sub New(ByVal lpInternalUseOnly As String)
      Try
        Cardinality = Cardinality.ecmSingleValued
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New()
      Try
        Cardinality = Cardinality.ecmSingleValued
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpCardinality As Cardinality)

      Me.New(lpMenuType, lpName, lpCardinality, True)

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpCardinality As Cardinality,
                   ByVal lpPersistent As Boolean)

      MyBase.New()
      With Me
        .Type = lpMenuType
        .Name = lpName
        .Cardinality = lpCardinality
        .mblnPersistent = lpPersistent
      End With

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValue As Object)
      Me.New(lpMenuType, lpName, lpValue, True)
    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpValue As Object,
               ByVal lpPersistent As Boolean)

      MyBase.New()
      With Me
        .Type = lpMenuType
        .Name = lpName
        .Cardinality = Cardinality.ecmSingleValued
        .Value = lpValue
        .mblnPersistent = lpPersistent
      End With

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpValues As Values)
      Me.New(lpMenuType, lpName, lpValues, True)
    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValues As Values,
                   ByVal lpPersistent As Boolean)

      MyBase.New()
      With Me
        .Type = lpMenuType
        .Name = lpName
        .Cardinality = Cardinality.ecmMultiValued
        .Values = lpValues
        .mblnPersistent = lpPersistent
      End With

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpCardinality As Cardinality,
               ByVal lpDefaultValue As Object)

      Me.New(lpMenuType, lpName, lpCardinality, lpDefaultValue, True)

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpCardinality As Cardinality,
                   ByVal lpDefaultValue As Object,
                   ByVal lpPersistent As Boolean)
      MyBase.New()

      With Me
        .Type = lpMenuType
        .Name = lpName
        .Cardinality = lpCardinality
        .DefaultValue = lpDefaultValue
        .mblnPersistent = lpPersistent
      End With

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpValue As Object,
               ByVal lpDefaultValue As Object)

      Me.New(lpMenuType, lpName, lpValue, lpDefaultValue, True)

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValue As Object,
                   ByVal lpDefaultValue As Object,
                   ByVal lpPersistent As Boolean)

      Me.New(lpMenuType, lpName, Core.Cardinality.ecmSingleValued, lpValue, lpDefaultValue, lpPersistent)

    End Sub

    <Obsolete("Direct construction of the ECMProperty class has been deprecated, please use PropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
               ByVal lpName As String,
               ByVal lpCardinality As Cardinality,
               ByVal lpValue As Object,
               ByVal lpDefaultValue As Object,
               ByVal lpPersistent As Boolean)

      MyBase.New()
      With Me
        .Type = lpMenuType
        .Name = lpName
        .Cardinality = lpCardinality
        .Value = lpValue
        .DefaultValue = lpDefaultValue
        .mblnPersistent = lpPersistent
      End With

    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Clears current value(s) from the property.
    ''' </summary>
    ''' <remarks>
    ''' Clears both the Value and Values properties
    ''' </remarks>
    Public Sub Clear() Implements IProperty.Clear
      Try
        mobjValue = Nothing
        mobjValues = New Values
        mobjValues.SetParentProperty(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Renames the property using the supplied information.
    ''' </summary>
    ''' <param name="lpNewName">
    ''' The new name for the property.  A check will be made to 
    ''' see if the original system name was the same as the 
    ''' original name.  If so then the new name will be applied 
    ''' to the system name as well.
    ''' </param>
    ''' <remarks>
    ''' If the original ID is the same as the original name 
    ''' then the new name will be applied to the ID as well.
    ''' </remarks>
    Public Sub Rename(ByVal lpNewName As String)
      Try
        Rename(lpNewName, Nothing)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Renames the property using the supplied information.
    ''' </summary>
    ''' <param name="lpNewName">
    ''' The new name for the property.
    ''' </param>
    ''' <param name="lpNewSystemName">
    ''' The new system name for the property.  If no value is 
    ''' specified then a check will be made to see if the original 
    ''' system name was the same as the original name.  If so then 
    ''' the new name will be applied to the system name as well.
    ''' </param>
    ''' <remarks>
    ''' If the original ID is the same as the original name 
    ''' then the new name will be applied to the ID as well.
    ''' </remarks>
    Public Sub Rename(ByVal lpNewName As String, ByVal lpNewSystemName As String)
      Try
        Dim lstrOriginalName As String = Name
        Name = lpNewName

        If Not String.IsNullOrEmpty(lpNewSystemName) Then
          SystemName = lpNewSystemName
        ElseIf String.Equals(SystemName, lstrOriginalName) Then
          SystemName = lpNewName
        ElseIf String.IsNullOrEmpty(SystemName) Then
          SystemName = lpNewName
        End If

        ' If the original ID was equal to the original name, let's change the ID as well.
        If String.Equals(ID, lstrOriginalName) = True Then
          ID = lpNewName
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub SetID(ByVal lpId As String)
      mstrID = lpId
    End Sub

    Public Sub SetPackedName(ByVal lpPackedName As String)
      mstrPackedName = lpPackedName
    End Sub

    Public Overrides Function ToString() As String
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function ToDebugString() As String Implements IProperty.ToDebugString
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Just revert to the default behavior
        Return Me.GetType.Name
      End Try
    End Function

#End Region

#Region "Friend Methods"

    Friend Sub SetPersistence(ByVal lpPersistent As Boolean)
      Try
        mblnPersistent = lpPersistent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Sub OnPropertyChanged(ByVal lpPropertyName As String)
      Try
        Me.OnPropertyChanged(New PropertyChangedEventArgs(lpPropertyName))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Overridable Sub OnPropertyChanged(ByVal e As PropertyChangedEventArgs)
      Try
        Dim lobjEventHandlers As PropertyChangedEventHandler = Me.PropertyChangedEvent
        If lobjEventHandlers IsNot Nothing Then
          lobjEventHandlers(Me, e)
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Private Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try

        If disposedValue = True Then
          Return "Property Disposed"
        End If

        Dim lobjReturnBuilder As New StringBuilder

        lobjReturnBuilder.AppendFormat("{0}: {1} - ", Me.GetType.Name, Name)

        lobjReturnBuilder.AppendFormat("{0} - ", EnumDescriptionTypeConverter.GetDescription(GetType(PropertyType), Type.ToString)) ' Type.ToString)

        Select Case Cardinality
          Case Core.Cardinality.ecmSingleValued
            If Value Is Nothing Then
              'lobjReturnBuilder.AppendFormat("{0}: {1}", Name, "Null")
              lobjReturnBuilder.Append("Null")
            ElseIf Value.ToString = String.Empty Then
              'lobjReturnBuilder.AppendFormat("{0}: {1}", Name, "Value not set")
              lobjReturnBuilder.Append("Value not set")
            End If
            lobjReturnBuilder.AppendFormat("{0}", Value.ToString)
          Case Else
            If Values Is Nothing OrElse Values.Count = 0 Then
              lobjReturnBuilder.Append("No Values")
            Else
              'Dim lobjValueBuilder As New Text.StringBuilder
              'For Each lobjValue As Value In Values
              For Each lobjValue As Object In Values
                If TypeOf (lobjValue) Is Value Then
                  lobjReturnBuilder.AppendFormat("{0}; ", lobjValue.Value.ToString)
                Else
                  lobjReturnBuilder.AppendFormat("{0}; ", lobjValue.ToString)
                End If
              Next
              lobjReturnBuilder = lobjReturnBuilder.Remove(lobjReturnBuilder.Length - 2, 2)
              'lobjReturnBuilder.AppendFormat("{0}: {1}", Name, lobjReturnBuilder.ToString)
            End If
        End Select

        Return lobjReturnBuilder.ToString

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Shared Function GetPackedName(ByVal lpInputName As String) As String
      Try
        lpInputName = lpInputName.Replace(" ", "")
        lpInputName = lpInputName.Replace("/", "")
        lpInputName = lpInputName.Replace("\", "")

        Return lpInputName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Determines whether or not the property has a value specified
    ''' </summary>
    ''' <returns>True if there is a value for the single valued case or at least one value for the multi-valued case</returns>
    ''' <remarks></remarks>
    Protected Overridable Function PropertyHasValue() As Boolean

      Try
        Select Case Me.Cardinality
          Case Core.Cardinality.ecmSingleValued
            If Value Is Nothing Then
              Return False
            End If
            If Value.ToString.Length = 0 Then
              Return False
            End If
          Case Core.Cardinality.ecmMultiValued
            If Values Is Nothing Then
              Return False
            End If
            If Values.Count = 0 Then
              Return False
            End If
            For Each lobjValue As Object In Values
              If TypeOf lobjValue Is Value Then
                If lobjValue.Value Is Nothing Then
                  Return False
                End If
              ElseIf TypeOf lobjValue Is String Then
                If String.IsNullOrEmpty(lobjValue) Then
                  Return False
                End If
              End If
            Next
        End Select

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod, Me.Name)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "IJsonSerializable(Of ECMProperty)"

    Public Overridable Function ToJson() As String Implements IJsonSerializable(Of ECMProperty).ToJson
      Try
        Return JsonConvert.SerializeObject(Me, Newtonsoft.Json.Formatting.Indented, DefaultJsonSerializerSettings.Settings)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function FromJson(lpJson As String) As ECMProperty Implements IJsonSerializable(Of ECMProperty).FromJson
      Try
        Return JsonConvert.DeserializeObject(lpJson, GetType(ECMProperty), DefaultJsonSerializerSettings.Settings)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IXmlSerializable Implementation"

    Public Function GetSchema() As System.Xml.Schema.XmlSchema Implements System.Xml.Serialization.IXmlSerializable.GetSchema
      Try
        Return Nothing
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Sub ReadXml(ByVal reader As System.Xml.XmlReader) Implements System.Xml.Serialization.IXmlSerializable.ReadXml
      Try

        If reader.IsEmptyElement Then
          reader.Read()
          Exit Sub
        End If

        ' We want to know the ECMProperty subclass
        ' The next line assumes that the type is the first attribute.
        ' We have to check based on ordinal because the attribute may have one of many labels including 'xsi:type', 'd3p1:xsi', 'd5p1:xsi'

        mstrXsiType = reader.GetAttribute(0)

        Dim lstrCurrentElementName As String = String.Empty

        ID = reader.GetAttribute("ID")
        Description = reader.GetAttribute("Description")

        Do Until reader.NodeType = XmlNodeType.EndElement AndAlso reader.Name = "ECMProperty"
          If reader.NodeType = XmlNodeType.Element Then
            lstrCurrentElementName = reader.Name
          ElseIf reader.NodeType = XmlNodeType.Text Then

            Select Case lstrCurrentElementName
              Case "Type"
                Type = [Enum].Parse(Type.GetType, reader.Value, True)
              Case "Cardinality"
                Cardinality = [Enum].Parse(Cardinality.GetType, reader.Value, True)
              Case "Name"
                Name = reader.Value
              Case "SystemName"
                SystemName = reader.Value
              Case "Value"
                If Not String.IsNullOrEmpty(reader.Value) Then
                  If Cardinality = Core.Cardinality.ecmSingleValued Then
                    Select Case Type
                      Case PropertyType.ecmDate
                        If TypeOf reader Is XmlStreamReader Then
                          If Not String.IsNullOrEmpty(DirectCast(reader, XmlStreamReader).LocaleString) Then
                            ' Try to parse the date based on the locale given (if available)
                            Dim ldatDateValue As DateTime = DateTime.MinValue
                            Dim lblnDateParseSuccess As Boolean = DateTime.TryParse(reader.Value,
                                                                    DirectCast(reader, XmlStreamReader).Culture,
                                                                    DateTimeStyles.None, ldatDateValue)
                            If lblnDateParseSuccess Then
                              Value = ldatDateValue
                            Else
                              Value = reader.Value
                            End If
                          End If
                        Else
                          Value = reader.Value
                        End If
                      Case Else
                        Value = reader.Value
                    End Select
                  ElseIf Cardinality = Core.Cardinality.ecmMultiValued Then
                    Values.Add(reader.Value)
                  End If
                End If
              Case "Values"
                ' TODO: Implement this...
              Case "anyType"
                If reader.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(reader.Value) Then
                  Select Case Type
                    Case PropertyType.ecmString
                      Values.AddString(reader.Value)
                    Case PropertyType.ecmBinary
                      Values.AddObject(reader.Value)
                    Case PropertyType.ecmBoolean
                      Values.AddBoolean(Boolean.Parse(reader.Value))
                    Case PropertyType.ecmDate
                      Values.AddDate(Date.Parse(reader.Value))
                    Case PropertyType.ecmDouble
                      Values.AddDouble(Double.Parse(reader.Value))
                    Case PropertyType.ecmGuid
                      Values.Add(reader.Value)
                    Case PropertyType.ecmHtml
                      Values.Add(reader.Value)
                    Case PropertyType.ecmLong
                      Values.AddLong(Long.Parse(reader.Value))
                    Case PropertyType.ecmObject
                      Values.Add(reader.Value)
                    Case PropertyType.ecmUri
                      Values.Add(reader.Value)
                    Case PropertyType.ecmXml
                      Values.Add(reader.Value)
                  End Select
                End If
            End Select
          End If
          reader.Read()
          Do Until reader.NodeType <> XmlNodeType.Whitespace
            reader.Read()
          Loop
        Loop

        reader.ReadEndElement()

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Overridable Sub WriteXml(ByVal writer As System.Xml.XmlWriter) Implements System.Xml.Serialization.IXmlSerializable.WriteXml
      Try

        With writer

          'If TypeOf Me Is SingletonProperty OrElse TypeOf Me Is MultiValueProperty Then
          '  ' This is not the base ECMProperty
          '  ' Write the xsi:type attribute
          '  .WriteAttributeString("xsi", "type", Me.GetType.Name)
          '  '.WriteAttributeString("xsi:type", Me.GetType.Name)
          'End If

          Dim lstrTypeName As String = Me.GetType.Name

          .WriteAttributeString("xsi", "type", lstrTypeName)

          WriteStandardElements(writer)

          If Cardinality = Core.Cardinality.ecmSingleValued Then
            ' Write the Value Element
            If Value IsNot Nothing Then
              Select Case Type
                Case PropertyType.ecmDate
                  If TypeOf Value Is DateTime AndAlso TypeOf writer Is EnhancedXmlTextWriter Then
                    '.WriteElementString("Value", Value.ToString())
                    .WriteElementString("Value", DirectCast(Value, DateTime).ToString(DirectCast(writer, EnhancedXmlTextWriter).Culture))
                  Else
                    .WriteElementString("Value", Value.ToString)
                  End If
                Case Else
                  .WriteElementString("Value", Value.ToString)
              End Select
            Else
              .WriteElementString("Value", Nothing)
            End If

          ElseIf Cardinality = Core.Cardinality.ecmMultiValued Then
            ' Write the Values Element
            If Values Is Nothing OrElse Values.Count = 0 Then
              .WriteElementString("Values", Nothing)
            Else
              ' Open the Values Element
              .WriteStartElement("Values")
              For Each lobjValue As Object In Values
                .WriteStartElement("anyType")
                '.WriteAttributeString("xsi:type", "xsd:string")
                .WriteAttributeString("xsi", "type", Me.GetType.Name)
                ' Check whether or not we are wrapped in a Value object.
                If TypeOf lobjValue Is Value Then
                  .WriteValue(DirectCast(lobjValue, Value).Value.ToString)
                Else
                  .WriteValue(lobjValue.ToString)
                End If
                .WriteEndElement()
              Next
              ' Close the Values element
              .WriteEndElement()
            End If
          End If

        End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Overridable Sub WriteStandardElements(ByVal writer As System.Xml.XmlWriter)
      Try

        With writer

          ' Write the ID attribute
          .WriteAttributeString("ID", ID)

          ' Write the Description attribute
          .WriteAttributeString("Description", Description)

          ' Write the Type element
          .WriteElementString("Type", Type.ToString)

          ' Write the Cardinality element
          .WriteElementString("Cardinality", Cardinality.ToString)

          ' Write the Name element
          .WriteElementString("Name", Name)

          ' Write the SystemName element
          .WriteElementString("SystemName", SystemName)

        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
      If Not Me.disposedValue Then
        If disposing Then
          ' DISPOSETODO: dispose managed state (managed objects).
          mobjValue = Nothing
          mobjValues = Nothing
          mstrName = Nothing
          mstrPackedName = Nothing
          mstrSystemName = Nothing
          mstrID = Nothing
          mobjDefaultValue = Nothing
          mstrDescription = Nothing
          mstrXsiType = Nothing
          menuCardinalityType = Nothing
          menuType = Nothing
          mblnPersistent = Nothing
        End If

        ' DISPOSETODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
        ' DISPOSETODO: set large fields to null.
      End If
      Me.disposedValue = True
    End Sub

    ' DISPOSETODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
      Dispose(True)
      GC.SuppressFinalize(Me)
    End Sub
#End Region

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
  End Class

End Namespace