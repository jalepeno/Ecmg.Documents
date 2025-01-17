'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Reflection
Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core
  ''' <summary>Property description object return in IClassification methods.</summary>

  <DebuggerDisplay("{DebuggerIdentifier(),nq}"),
  XmlInclude(GetType(ClassificationBinaryProperty)),
  XmlInclude(GetType(ClassificationBooleanProperty)),
  XmlInclude(GetType(ClassificationDateTimeProperty)),
  XmlInclude(GetType(ClassificationDoubleProperty)),
  XmlInclude(GetType(ClassificationGuidProperty)),
  XmlInclude(GetType(ClassificationHtmlProperty)),
  XmlInclude(GetType(ClassificationLongProperty)),
  XmlInclude(GetType(ClassificationObjectProperty)),
  XmlInclude(GetType(ClassificationStringProperty)),
  XmlInclude(GetType(ClassificationUriProperty)),
  XmlInclude(GetType(ClassificationXmlProperty))>
  Public Class ClassificationProperty
    Inherits ECMProperty
    Implements ISingletonProperty
    Implements IXmlSerializable

#Region "Class Enumerations"

    ''' <summary>
    ''' Specifies a PropertySettability constant, 
    ''' which indicates when the value of a property can be set.
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum SettabilityEnum

      ''' <summary>
      ''' Indicates that a property is read-only; 
      ''' only the server can set its value.
      ''' </summary>
      ''' <remarks></remarks>
      READ_ONLY = 3

      ''' <summary>
      ''' Indicates that a property is read/write; 
      ''' you can set its value at any time.
      ''' </summary>
      ''' <remarks></remarks>
      READ_WRITE = 0

      ''' <summary>
      ''' Indicates that you can only set the value of a 
      ''' property before you check in the object to which it belongs.
      ''' </summary>
      ''' <remarks></remarks>
      SETTABLE_ONLY_BEFORE_CHECKIN = 1

      ''' <summary>
      ''' Indicates that you can only set the value of a property 
      ''' when you create the object to which it belongs. 
      ''' Once you save the object for the first time, 
      ''' the property's value cannot be changed.
      ''' </summary>
      ''' <remarks></remarks>
      SETTABLE_ONLY_ON_CREATE = 2

    End Enum

#End Region

#Region "Class Variables"

    Private mblnSearchable As Boolean = True
    Private mblnSelectable As Boolean = True
    Private mblnIsSystemProperty As Boolean
    Private mblnIsHidden As Boolean
    Private mblnIsRequired As Boolean
    Private mblnIsInherited As Boolean
    Private mobjSubscribedClasses As DocumentClasses
    Private mobjSubscribedClassNames As New Generic.List(Of String)
    Private menuSettability As SettabilityEnum = SettabilityEnum.READ_WRITE
    Private mobjChoiceList As ChoiceLists.ChoiceList = Nothing
    Private mstrChoiceListName As String = String.Empty

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Determines whether or not this property association to the document class is inherited from a parent document class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsInherited() As Boolean
      Get
        Return mblnIsInherited
      End Get
      Set(ByVal value As Boolean)
        mblnIsInherited = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not this property is required for the associated document class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsRequired() As Boolean
      Get
        Return mblnIsRequired
      End Get
      Set(ByVal value As Boolean)
        mblnIsRequired = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not this property is hidden in this document class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsHidden() As Boolean
      Get
        Return mblnIsHidden
      End Get
      Set(ByVal value As Boolean)
        mblnIsHidden = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether this property is system defined or customer defined.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IsSystemProperty() As Boolean
      Get
        Return mblnIsSystemProperty
      End Get
      Set(ByVal value As Boolean)
        mblnIsSystemProperty = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not this property may be used as a search criteria.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Searchable() As Boolean
      Get
        Return mblnSearchable
      End Get
      Set(ByVal value As Boolean)
        mblnSearchable = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies how the property value may or may not be set
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Settability() As SettabilityEnum
      Get
        Return menuSettability
      End Get
      Set(ByVal value As SettabilityEnum)
        menuSettability = value
      End Set
    End Property

    ''' <summary>
    ''' Specifies whether or not this property may be returned as a result column in a search.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Selectable() As Boolean
      Get
        Return mblnSelectable
      End Get
      Set(ByVal value As Boolean)
        mblnSelectable = value
      End Set
    End Property

    Public Property ChoiceList() As ChoiceLists.ChoiceList
      Get
        Return mobjChoiceList
      End Get
      Set(ByVal value As ChoiceLists.ChoiceList)
        mobjChoiceList = value
        mstrChoiceListName = mobjChoiceList.Name
      End Set
    End Property

    Public ReadOnly Property ChoiceListName As String
      Get
        Return mstrChoiceListName
      End Get
    End Property

    ''' <summary>
    ''' Lists all of the document classes to which this property is associated.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SubscribedClasses() As DocumentClasses
      Get
        Return mobjSubscribedClasses
      End Get
      Set(ByVal value As DocumentClasses)
        Try
          mobjSubscribedClasses = value
          SubscribedClassNames = GetSubscribedClassNames(value)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Lists all of the document class names to which this property is associated.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SubscribedClassNames() As Generic.List(Of String)
      Get
        Return mobjSubscribedClassNames
      End Get
      Set(ByVal value As Generic.List(Of String))
        mobjSubscribedClassNames = value
      End Set
    End Property

    Public Shadows Property Value As Object Implements ISingletonProperty.Value
      Get
        Try
          Return MyBase.Value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Object)
        Try
          MyBase.Value = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    <Obsolete("Direct invocation of constructor has been deprecated, please use ClassificationPropertyFactory.Create instead.")>
    Public Sub New()
      MyBase.New()
    End Sub

    Protected Sub New(ByVal lpType As PropertyType)
      MyBase.New(Nothing)
      Try
        Type = lpType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub New(ByVal lpType As PropertyType, ByVal lpName As String)
      MyBase.New(Nothing)
      Try
        Type = lpType
        Name = lpName
        SystemName = lpName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Sub New(ByVal lpType As PropertyType, ByVal lpName As String, ByVal lpSystemName As String)
      MyBase.New(Nothing)
      Try
        Type = lpType
        Name = lpName
        SystemName = lpSystemName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    <Obsolete("Direct invocation of constructor has been deprecated, please use ClassificationPropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpCardinality As Cardinality,
                   Optional ByVal lpSearchable As Boolean = True,
                   Optional ByVal lpSelectable As Boolean = True,
                   Optional ByVal lpSystemProperty As Boolean = True,
                   Optional ByVal lpIsHidden As Boolean = False,
                   Optional ByVal lpIsRequired As Boolean = False,
                   Optional ByVal lpIsInherited As Boolean = False,
                   Optional ByVal lpSystemName As String = "")

      MyBase.New(lpMenuType, lpName, lpCardinality)
      Try
        Searchable = lpSearchable
        Selectable = lpSelectable
        IsSystemProperty = lpSystemProperty
        IsHidden = lpIsHidden
        IsRequired = lpIsRequired
        IsInherited = lpIsInherited
        If lpSystemName.Length > 0 Then
          SystemName = lpSystemName
        Else
          SystemName = lpName
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    <Obsolete("Direct invocation of constructor has been deprecated, please use ClassificationPropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValue As Object,
                   Optional ByVal lpSearchable As Boolean = True,
                   Optional ByVal lpSelectable As Boolean = True,
                   Optional ByVal lpSystemProperty As Boolean = True,
                   Optional ByVal lpIsHidden As Boolean = False,
                   Optional ByVal lpIsRequired As Boolean = False,
                   Optional ByVal lpIsInherited As Boolean = False,
                   Optional ByVal lpSystemName As String = "")

      MyBase.New(lpMenuType, lpName, lpValue)
      Try
        Searchable = lpSearchable
        Selectable = lpSelectable
        IsSystemProperty = lpSystemProperty
        IsHidden = lpIsHidden
        IsRequired = lpIsRequired
        IsInherited = lpIsInherited
        If lpSystemName.Length > 0 Then
          SystemName = lpSystemName
        Else
          SystemName = lpName
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    <Obsolete("Direct invocation of constructor has been deprecated, please use ClassificationPropertyFactory.Create instead.")>
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValues As Core.Values,
                   Optional ByVal lpSearchable As Boolean = True,
                   Optional ByVal lpSelectable As Boolean = True,
                   Optional ByVal lpSystemProperty As Boolean = True,
                   Optional ByVal lpIsHidden As Boolean = False,
                   Optional ByVal lpIsRequired As Boolean = False,
                   Optional ByVal lpIsInherited As Boolean = False,
                   Optional ByVal lpSystemName As String = "")

      MyBase.New(lpMenuType, lpName, lpValues)
      Try
        Searchable = lpSearchable
        Selectable = lpSelectable
        IsSystemProperty = lpSystemProperty
        IsHidden = lpIsHidden
        IsRequired = lpIsRequired
        IsInherited = lpIsInherited
        If lpSystemName.Length > 0 Then
          SystemName = lpSystemName
        Else
          SystemName = lpName
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    ''' <summary>
    ''' Constructs from a basic ECMProperty
    ''' </summary>
    ''' <param name="lpProperty"></param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal lpProperty As ECMProperty)
      MyBase.New(Nothing)

      Type = lpProperty.Type
      Name = lpProperty.Name
      SystemName = lpProperty.SystemName
      Cardinality = lpProperty.Cardinality
      Searchable = True
      Selectable = True
      IsSystemProperty = False
      IsHidden = False
      IsRequired = False
      IsInherited = False
      If String.IsNullOrEmpty(SystemName) AndAlso Not String.IsNullOrEmpty(lpProperty.PackedName) Then
        SystemName = lpProperty.PackedName
      End If

    End Sub

    ''' <summary>
    ''' Constructs from a generic IProperty
    ''' </summary>
    ''' <param name="lpProperty"></param>
    ''' <remarks></remarks>
    Friend Sub New(ByVal lpProperty As IProperty)
      MyBase.New(Nothing)

      Type = lpProperty.Type
      Name = lpProperty.Name
      SystemName = lpProperty.SystemName
      Cardinality = lpProperty.Cardinality
      Searchable = True
      Selectable = True
      IsSystemProperty = False
      IsHidden = False
      IsRequired = False
      IsInherited = False

    End Sub

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Converts the current object to a base ECMProperty object.
    ''' </summary>
    ''' <returns>A base ECMProperty object based on the current object.</returns>
    ''' <remarks>If there are values present, they will be retained.</remarks>
    Public Function ToECMProperty() As ECMProperty

      Try
        Return ToECMProperty(True)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

    ''' <summary>
    ''' Converts the current object to a base ECMProperty object.
    ''' </summary>
    ''' <param name="lpRetainValue">
    ''' Specifies whether or not to retain any values present in the return value.
    ''' </param>
    ''' <returns>
    ''' A base ECMProperty object based on the current object.
    ''' </returns>
    ''' <remarks>
    ''' Depending on the value specified for lpRetainValue the 
    ''' present value(s) may or may not be retained in the output object.
    ''' </remarks>
    Public Function ToECMProperty(ByVal lpRetainValue As Boolean) As ECMProperty

      Try
        Dim lobjECMProperty As ECMProperty = PropertyFactory.Create(Me.Type, Me.Name, Me.SystemName, Me.Cardinality)

        With lobjECMProperty
          .SetID(Me.ID)
          '.Type = Me.Type
          '.Name = Me.Name
          '.Cardinality = Me.Cardinality
          If lpRetainValue Then
            If .Cardinality = Cardinality.ecmSingleValued Then
              .Value = Me.Value
            Else
              .Values = Me.Values
            End If
          End If
        End With

        Return lobjECMProperty

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

#Region "Friend Methods"

    Friend Function GetSubscribedClasses(lpDocumentClasses As DocumentClasses) As DocumentClasses
      Try

        Dim lobjSubscribedClasses As New DocumentClasses

        For Each lobjDocumentClass As DocumentClass In lpDocumentClasses
          If lobjDocumentClass.IsSubscribed(Me) Then
            lobjSubscribedClasses.Add(lobjDocumentClass)
          End If
        Next

        Return lobjSubscribedClasses

      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Private Function GetSubscribedClassNames(lpDocumentClasses As DocumentClasses) As List(Of String)
      Try
        Dim lobjClassNames As New List(Of String)

        For Each lobjDocumentClass As DocumentClass In lpDocumentClasses
          lobjClassNames.Add(lobjDocumentClass.Name)
        Next
        Return lobjClassNames
      Catch ex As Exception
        ApplicationLogging.LogException(ex, MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Returns a string representation of the current object.
    ''' </summary>
    ''' <returns>A string representation of the current object.</returns>
    ''' <remarks></remarks>
    Protected Overrides Function DebuggerIdentifier() As String
      Try
        If DefaultValue IsNot Nothing Then
          If TypeOf (DefaultValue) Is String AndAlso DefaultValue.ToString.Length = 0 Then
            Return String.Format("{0} {1}: {2} <{3}>",
                     Cardinality.ToString.Substring(3), Type.ToString.Substring(3),
                     Name, Settability.ToString)
          Else
            Return String.Format("{0} {1}: {2}; DefaultValue: {3} <{4}>",
                     Cardinality.ToString.Substring(3), Type.ToString.Substring(3),
                     Name, DefaultValue.ToString, Settability.ToString)
          End If
        Else
          Return String.Format("{0} {1}: {2} <{3}>",
                   Cardinality.ToString.Substring(3), Type.ToString.Substring(3), Name,
                   Settability.ToString)
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Private Sub SetChoiceListName(lpName As String)
      Try
        mstrChoiceListName = lpName
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Shared Sub SetDefaultValues(ByVal lpProperty As ClassificationProperty)
      Try

        With lpProperty
          .Searchable = True
          .Selectable = True
          .IsSystemProperty = False
          .IsHidden = False
          .IsRequired = False
          .IsInherited = False
          .SystemName = .PackedName
        End With

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "IXmlSerializable Implementation"

    Protected Overrides Sub WriteStandardElements(ByVal writer As System.Xml.XmlWriter)
      Try

        MyBase.WriteStandardElements(writer)

        With writer

          ' Write the IsInherited element
          .WriteElementString("IsInherited", IsInherited)

          ' Write the IsRequired element
          .WriteElementString("IsRequired", IsRequired)

          ' Write the IsHidden element
          .WriteElementString("IsHidden", IsHidden)

          ' Write the IsSystemProperty element
          .WriteElementString("IsSystemProperty", IsSystemProperty)

          ' Write the Searchable element
          .WriteElementString("Searchable", Searchable)

          ' Write the Settability element
          .WriteElementString("Settability", Settability.ToString)

          ' Write the Selectable element
          .WriteElementString("Selectable", Selectable)

          ' TODO: Add SubscribedClasses

        End With
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "ICloneable Implementation"

    Public Overloads Function Clone() As Object
      Try
        Return ClassificationPropertyFactory.Create(Me.Type, Me.Name, Me.SystemName, Me.Cardinality, Me.DefaultValue)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace