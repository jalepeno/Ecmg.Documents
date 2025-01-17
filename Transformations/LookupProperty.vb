'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Transformations

  <Serializable()>
  Public Class LookupProperty

#Region "Class Constants"

    Friend Const PARAM_PROPERTY_NAME As String = "PropertyName"
    Friend Const PARAM_PROPERTY_SCOPE As String = "PropertyScope"
    Friend Const PARAM_PROPERTY_TYPE As String = "PropertyType"
    Friend Const PARAM_VERSION_INDEX As String = "VersionIndex"
    Friend Const PARAM_VALUE_INDEX As String = "ValueIndex"
    Friend Const PARAM_AUTO_CREATE As String = "AutoCreate"
    Friend Const PARAM_PERSISTENT As String = "Persistent"

#End Region

#Region "Class Variables"

    Private mstrPropertyName As String = String.Empty
    Private menuPropertyScope As PropertyScope
    Private mintVersionIndex As Integer
    Private menuValueIndex As Values.ValueIndexEnum = Values.ValueIndexEnum.First
    Private mblnAutoCreate As Boolean
    Private mblnPersistent As Boolean
    Private menuPropertyType As PropertyType = PropertyType.ecmString

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets or sets the LookupProperty name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Interchangeable with PropertyName</remarks>
    <XmlIgnore()>
    Public Property Name() As String
      Get
        Return mstrPropertyName
      End Get
      Set(ByVal value As String)
        mstrPropertyName = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the LookupProperty name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Interchangeable with Name</remarks>
    Public Property PropertyName() As String
      Get
        Return mstrPropertyName
      End Get
      Set(ByVal value As String)
        mstrPropertyName = value
      End Set
    End Property

    Public Property PropertyScope() As PropertyScope
      Get
        Return menuPropertyScope
      End Get
      Set(ByVal value As PropertyScope)
        menuPropertyScope = value
      End Set
    End Property

    Public Property Type As PropertyType
      Get
        Return menuPropertyType
      End Get
      Set(ByVal value As PropertyType)
        menuPropertyType = value
      End Set
    End Property

    Public Property VersionIndex() As Integer
      Get
        Return mintVersionIndex
      End Get
      Set(ByVal value As Integer)
        mintVersionIndex = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the index of the value to be used 
    ''' in the case of a multi-valued source property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ValueIndex() As Values.ValueIndexEnum
      Get
        Return menuValueIndex
      End Get
      Set(ByVal value As Values.ValueIndexEnum)
        menuValueIndex = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value that determines if the 
    ''' destination property should be created automatically
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <XmlAttribute()>
    Public Property AutoCreate() As Boolean
      Get
        Return mblnAutoCreate
      End Get
      Set(ByVal value As Boolean)
        mblnAutoCreate = value
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
    <XmlAttribute()>
    Public Overloads Property Persistent() As Boolean
      Get
        Return mblnPersistent
      End Get
      Set(ByVal value As Boolean)
        mblnPersistent = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpPropertyName As String,
                   ByVal lpPropertyScope As Core.PropertyScope,
                   Optional ByVal lpVersionIndex As Integer = 0)

      Try
        PropertyName = lpPropertyName
        PropertyScope = lpPropertyScope
        VersionIndex = lpVersionIndex
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

#End Region

#Region "Public Methods"

    Public Function GetParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_PROPERTY_NAME) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PROPERTY_NAME, PropertyName,
            "Specifies the name of the target property to modify."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_SCOPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PROPERTY_SCOPE, PropertyScope, GetType(PropertyScope),
            "Specifies whether the target property is a document or version property."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_TYPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PROPERTY_TYPE, Type, GetType(PropertyType),
            "Specifies the target property data type."))
        End If

        If lobjParameters.Contains(PARAM_VERSION_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_VERSION_INDEX, VersionIndex,
            "Specifies the version index."))
        End If

        If lobjParameters.Contains(PARAM_VALUE_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_VALUE_INDEX, ValueIndex, GetType(Values.ValueIndexEnum),
            "Specifies the target property data type."))
        End If

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Function GetProperty(ByVal lpMetaHolder As Core.IMetaHolder) As Core.ECMProperty
      Try
        Return DataLookup.GetProperty(Me, lpMetaHolder)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a property in the specified document using the current LookupProperty instance
    ''' </summary>
    ''' <param name="lpDocument">The document to create the property in</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateProperty(ByVal lpDocument As Document) As ECMProperty
      Try

        Dim lobjDocument As Document = Nothing
        Dim lblnCreateDocumentSuccess As Boolean

        lblnCreateDocumentSuccess = lpDocument.CreateProperty(Me.Name, , Me.Type, Me.PropertyScope)

        If lblnCreateDocumentSuccess = True Then
          Return GetProperty(lpDocument)
        Else
          Return Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

    ''' <summary>
    ''' Creates a property in the specified folder using the current LookupProperty instance
    ''' </summary>
    ''' <param name="lpFolder">The folder to create the property in</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateProperty(ByVal lpFolder As Folder) As ECMProperty
      Try

        Dim lblnCreateDocumentSuccess As Boolean

        lblnCreateDocumentSuccess = lpFolder.CreateProperty(Me.Name, Nothing, Me.Type, Cardinality.ecmSingleValued, True)

        If lblnCreateDocumentSuccess = True Then
          Return GetProperty(lpFolder)
        Else
          Return Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Friend Methods"

    Friend Shared Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_PROPERTY_NAME) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PROPERTY_NAME, String.Empty,
            "Specifies the name of the lookup property."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_SCOPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PROPERTY_SCOPE, PropertyScope.VersionProperty, GetType(PropertyScope),
            "Specifies what the scope is for the property, most typically this is either version or document."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_TYPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PROPERTY_TYPE, PropertyType.ecmString, GetType(PropertyType),
            "Specifies the target property data type."))
        End If

        If lobjParameters.Contains(PARAM_VERSION_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmLong, PARAM_VERSION_INDEX, -1,
            "Specifies which version(s) apply.  Zero indicates the first version, -1 indicates all versions."))
        End If

        If lobjParameters.Contains(PARAM_VALUE_INDEX) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_VALUE_INDEX, Values.ValueIndexEnum.First, GetType(Values.ValueIndexEnum),
            "For actions that read existing values from a multi-valued field, this specifies whether to read the first or last value."))
        End If

        If lobjParameters.Contains(PARAM_AUTO_CREATE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_AUTO_CREATE, False,
            "Specifies whether or not to automatically create this property if it does not exist.  NOTE: Auto creation will always create a single-valued string property."))
        End If

        If lobjParameters.Contains(PARAM_PERSISTENT) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_PERSISTENT, False,
            "For properties that are automatically created, this specifies whether they are persisted at the end of the transformation."))
        End If

        Return lobjParameters

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Overloaded Methods"

    Public Overrides Function ToString() As String

      Try
        Return String.Format("{0}~{1}", PropertyScope.ToString, mstrPropertyName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try

    End Function

#End Region

  End Class

End Namespace