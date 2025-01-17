'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Providers

  ''' <summary>
  ''' Used for defining a content source.  The collection of provider properties along with the provider itself defines a content source instance.
  ''' </summary>
  ''' <remarks></remarks>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class ProviderProperty
    Implements IComparable
    Implements IProperty
    Implements IEquatable(Of ProviderProperty)
    Implements INotifyPropertyChanged

#Region "Class Events"

    Event ProviderPropertyValueChanged As ProviderPropertyValueChangedEventHandler

#End Region

#Region "Class Variables"

    Private mstrPropertyName As String
    Private mobjPropertyValue As Object
    Private ReadOnly mobjPropertyType As Type
    Private mblnRequired As Boolean
    Private mblnVisible As Boolean = True
    Private mintSequenceNumber As Integer = -1
    Private mstrDescription As String = String.Empty
    Private mstrDisplayName As String = String.Empty
    Private mblnSupportsValueList As Boolean = False
    Private ReadOnly mblnIsKey As Boolean = True
    Private mobjParentProvider As IProvider
    Private ReadOnly menuDisplayType As DisplayType = DisplayType.Standard

#End Region

#Region "Public Events"

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Public Properties"

    Public ReadOnly Property ParentProvider As IProvider
      Get
        Try
          Return mobjParentProvider
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property PropertyName() As String
      Get
        Return mstrPropertyName
      End Get
      Set(ByVal value As String)
        Try
          mstrPropertyName = value
          OnPropertyChanged("PropertyName")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property PropertyValue() As Object
      Get
        Return mobjPropertyValue
      End Get
      Set(ByVal value As Object)
        Try
          ' Cache the old value
          Dim lobjOldValue As Object
          If TypeOf value IsNot Xml.XmlNode() Then
            lobjOldValue = mobjPropertyValue
            mobjPropertyValue = value
            If Not value.Equals(lobjOldValue) Then
              OnPropertyChanged("PropertyValue")
            End If
            RaiseEvent ProviderPropertyValueChanged(Me, New Arguments.ProviderPropertyValueChangedEventArgs(Me, lobjOldValue, value))
          Else
            ' For some reason we occasionally get xml node arrarys from the reader instead of the actual values
            ' Take the first value
            For Each lobjNode As Xml.XmlNode In value
              If TypeOf lobjNode Is Xml.XmlText Then
                lobjOldValue = mobjPropertyValue
                mobjPropertyValue = lobjNode.InnerText
                If Not value.Equals(lobjOldValue) Then
                  OnPropertyChanged("PropertyValue")
                End If
                RaiseEvent ProviderPropertyValueChanged(Me, New Arguments.ProviderPropertyValueChangedEventArgs(Me, lobjOldValue, mobjPropertyValue))
                Exit For
              Else
                Continue For
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

    Public ReadOnly Property PropertyType() As Type
      Get
        Return mobjPropertyType
      End Get
    End Property

    Public ReadOnly Property Required() As Boolean
      Get
        Return mblnRequired
      End Get
    End Property

    Public Overridable ReadOnly Property DisplayType() As DisplayType
      Get
        Return menuDisplayType
      End Get
    End Property

    Public Overridable ReadOnly Property IsKey() As Boolean
      Get
        Return mblnIsKey
      End Get
    End Property

    <XmlIgnore()>
    Public ReadOnly Property Visible() As Boolean
      Get
        Return mblnVisible
      End Get
    End Property

    Public ReadOnly Property SequenceNumber() As Integer
      Get
        Return mintSequenceNumber
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets a description for this property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property Description() As String Implements Core.IProperty.Description
      Get
        Return mstrDescription
      End Get
      Set(ByVal value As String)
        mstrDescription = value
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
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not this property should be able to get 
    ''' a list of available values for configuring an object store.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlIgnore()>
    Public Property SupportsValueList As Boolean
      Get
        Try
          If PropertyType IsNot Nothing AndAlso PropertyType.Name = "Boolean" Then
            mblnSupportsValueList = True
          End If
          Return mblnSupportsValueList
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try

      End Get
      Set(ByVal value As Boolean)
        Try
          mblnSupportsValueList = value
          OnPropertyChanged("SupportsValueList")
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

    End Sub

    Public Sub New(ByVal lpPropertyName As String, ByVal lpPropertyValue As String)
      Try
        mstrPropertyName = lpPropertyName
        mobjPropertyValue = lpPropertyValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpPropertyName As String, ByVal lpPropertyType As Type, ByVal lpRequired As Boolean, ByVal lpIsKey As Boolean)
      Me.New(lpPropertyName, lpPropertyType, lpRequired,,,,, lpIsKey)
    End Sub

    Public Sub New(ByVal lpPropertyName As String,
                  ByVal lpPropertyType As Type,
                  Optional ByVal lpRequired As Boolean = True,
                  Optional ByVal lpPropertyValue As String = "",
                  Optional ByVal lpSequenceNumber As Integer = -1,
                  Optional ByVal lpDescription As String = "",
                  Optional ByVal lpSupportsValueList As Boolean = False,
                  Optional ByVal lpIsKey As Boolean = False,
                  Optional ByVal lpDisplayType As DisplayType = DisplayType.Standard)
      Try
        mstrPropertyName = lpPropertyName
        mobjPropertyType = lpPropertyType
        mobjPropertyValue = lpPropertyValue
        mblnRequired = lpRequired
        mintSequenceNumber = lpSequenceNumber
        mstrDescription = lpDescription
        mblnSupportsValueList = lpSupportsValueList
        mblnIsKey = lpIsKey
        menuDisplayType = lpDisplayType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Gets a list of available values to set for this property.
    ''' </summary>
    ''' <param name="lpProvider">A reference to the provider that 
    ''' will get the list of available values</param>
    ''' <returns>An enumerable collection of strings.</returns>
    ''' <remarks></remarks>
    Public Overridable Function GetAvailableValues(ByVal lpProvider As IProvider) As IEnumerable(Of String)
      Try
        Return lpProvider.GetAvailableValues(Me)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function ReadyToGetAvailableValues() As Boolean
      Try

        If SupportsValueList = False Then
          Throw New Exceptions.PropertyDoesNotSupportValueListException(Me)
        End If

        If ParentProvider IsNot Nothing Then
          Return ParentProvider.ReadyToGetAvailableValues(Me)
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overridable Function ReadyToGetAvailableValues(ByVal lpProvider As IProvider) As Boolean
      Try

        If SupportsValueList = False Then
          Throw New Exceptions.PropertyDoesNotSupportValueListException(Me)
        End If

        Return lpProvider.ReadyToGetAvailableValues(Me)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
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

#Region "Public Overrides Methods"

    Public Overloads Function Equals(ByVal other As ProviderProperty) As Boolean Implements System.IEquatable(Of ProviderProperty).Equals
      Try

        If other Is Nothing Then
          Return False
        End If

        ' Compare the parts
        ' PropertyName
        If Me.PropertyName <> other.PropertyName Then
          Return False
        End If

        ' PropertyValue
        If (TypeOf Me.PropertyValue IsNot Object) AndAlso (TypeOf other.PropertyValue IsNot Object) Then
          If Me.PropertyValue <> other.PropertyValue Then
            Return False
          End If
        End If

        ' Required 
        If Me.Required <> other.Required Then
          Return False
        End If

        Return True

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Overrides Function Equals(ByVal obj As Object) As Boolean
      Try
        ' Make sure it is a real ProviderProperty
        If TypeOf (obj) Is ProviderProperty Then
          Return Equals(CType(obj, ProviderProperty))
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#End Region

#Region "Friend Methods"

    Friend Sub SetParentProvider(ByVal lpParentProvider As IProvider)
      Try
        mobjParentProvider = lpParentProvider
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Friend Sub SetRequired(ByVal lpRequired As Boolean)
      Try
        mblnRequired = lpRequired
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Friend Sub SetVisible(ByVal lpVisible As Boolean)
      Try
        mblnVisible = lpVisible
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Friend Sub SetSequenceNumber(ByVal lpSequenceNumber As Integer)
      Try
        mintSequenceNumber = lpSequenceNumber
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try

        If PropertyValue Is Nothing Then
          Return String.Format("{0} (Value is null)", PropertyName)
        ElseIf TypeOf (PropertyValue) Is String AndAlso DirectCast(PropertyValue, String).Length = 0 Then
          Return String.Format("{0}=(Value not set)", PropertyName)
        Else
          Return String.Format("{0}={1}", PropertyName, PropertyValue.ToString)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Sub OnPropertyChanged(ByVal lpPropertyName As String)
      Try
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(lpPropertyName))
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo

      If TypeOf obj Is ProviderProperty Then
        Return PropertyName.CompareTo(obj.PropertyName)
      Else
        Throw New ArgumentException("Object is not a ProviderProperty")
      End If

    End Function

#End Region

    Public Class SequenceNumberComparer
      Implements IComparer

      Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        ' Two null objects are equal
        If (x Is Nothing) And (y Is Nothing) Then Return 0

        ' Any non-null object is greater than a null object
        If (x Is Nothing) Then Return 1
        If (y Is Nothing) Then Return -1

        ' Cast both objects to roviderProperty, and do the comparison
        ' (Throws an exception if the objects aren't  ProviderProperty objects.)
        Dim propertyX As ProviderProperty = DirectCast(x, ProviderProperty)
        Dim propertyY As ProviderProperty = DirectCast(y, ProviderProperty)

        If propertyX.SequenceNumber < propertyY.SequenceNumber Then
          Return -1
        Else
          Return 1
        End If

      End Function

    End Class

    Public Overridable Property Cardinality As Core.Cardinality Implements Core.IProperty.Cardinality
      Get
        Return Core.Cardinality.ecmSingleValued
      End Get
      Set(value As Core.Cardinality)
        ' Ignore set operations
      End Set
    End Property

    Public Sub Clear() Implements Core.IProperty.Clear

    End Sub

    Public Property DefaultValue As Object Implements Core.IProperty.DefaultValue
      Get
        Return String.Empty
      End Get
      Set(value As Object)
        ' Ignore set operations
      End Set
    End Property

    Public ReadOnly Property HasValue As Boolean Implements Core.IProperty.HasValue
      Get
        If PropertyValue IsNot Nothing Then
          Return True
        Else
          Return False
        End If
      End Get
    End Property

    Public Overridable ReadOnly Property HasStandardValues As Boolean Implements IProperty.HasStandardValues
      Get
        Try
          Select Case Me.Type
            Case Core.PropertyType.ecmBoolean
              Return True
          End Select
          Return False
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
          Select Case Me.Type
            Case Core.PropertyType.ecmBoolean
              Return SingletonBooleanProperty.GetStandardValues()
          End Select
          Select Case Me.PropertyType.Name
            Case GetType(Boolean).Name
              Return SingletonBooleanProperty.GetStandardValues()
          End Select
          Return New List(Of String)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property Name As String Implements Core.IProperty.Name
      Get
        Return PropertyName
      End Get
      Set(value As String)
        Try
          PropertyName = value
          OnPropertyChanged("Name")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property Persistent As Boolean Implements Core.IProperty.Persistent
      Get
        Return True
      End Get
    End Property

    Public Property SystemName As String Implements Core.IProperty.SystemName
      Get
        Return PropertyName
      End Get
      Set(value As String)
        Try
          PropertyName = value
          OnPropertyChanged("SystemName")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Type As Core.PropertyType Implements Core.IProperty.Type
      Get
        Return Core.PropertyType.ecmString
      End Get
      Set(value As Core.PropertyType)
        ' Ignore set operations
      End Set
    End Property

    Public Property Value As Object Implements Core.IProperty.Value
      Get
        Return PropertyValue
      End Get
      Set(value As Object)
        Try
          PropertyValue = value
          OnPropertyChanged("Value")
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Property Values As Object Implements Core.IProperty.Values
      Get
        'Throw New NotImplementedException
        ' Return PropertyValue
        Return Nothing
      End Get
      Set(value As Object)
        'Throw New NotImplementedException
        ' PropertyValue = value
        ' Ignore set operation since all provider properties should be single-valued.
      End Set
    End Property

  End Class

End Namespace