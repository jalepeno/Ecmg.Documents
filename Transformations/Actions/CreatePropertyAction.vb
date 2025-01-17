'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Scripting
Imports Documents.Utilities

#End Region

Namespace Transformations
  ''' <summary>Action used to create a new property.</summary>
  <Serializable()>
  Public Class CreatePropertyAction
    Inherits PropertyAction

#Region "Class Constants"

    Private Const ACTION_NAME As String = "CreateProperty"
    Friend Const PARAM_CARDINALITY As String = "Cardinality"
    Friend Const PARAM_FIRST_VERSION_ONLY As String = "FirstVersionOnly"
    Friend Const PARAM_PERSISTENT As String = "Persistent"
    Friend Const PARAM_PROPERTY_TYPE As String = "PropertyType"
    Friend Const PARAM_PROPERTY_VALUE As String = "PropertyValue"
    Friend Const PARAM_VERSION_SCOPE As String = "VersionScope"

#End Region

#Region "Class Variables"

    Private mstrPropertyValue As Object
    Private menuPropertyType As Core.PropertyType
    Private mblnFirstVersionOnly As Boolean
    Private menuVersionScope As VersionScope = Transformations.VersionScope.AllVersions
    Private mblnPersistent As Boolean = True
    Private menuCardinality As Core.Cardinality = Core.Cardinality.ecmSingleValued

#End Region

#Region "Public Properties"

    Public Overrides ReadOnly Property ActionName As String
      Get
        Return ACTION_NAME
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the default value for the property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PropertyValue() As Object
      Get
        Return mstrPropertyValue
      End Get
      Set(ByVal Value As Object)
        mstrPropertyValue = Value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the data type of the property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PropertyType() As Core.PropertyType
      Get
        Return menuPropertyType
      End Get
      Set(ByVal Value As Core.PropertyType)
        menuPropertyType = Value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the cardinality of the property
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Cardinality() As Core.Cardinality
      Get
        Return menuCardinality
      End Get
      Set(ByVal Value As Core.Cardinality)
        menuCardinality = Value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value which determines if the property 
    ''' should be created for the first version of the document only
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FirstVersionOnly() As Boolean
      Get
        Return mblnFirstVersionOnly
      End Get
      Set(ByVal Value As Boolean)
        mblnFirstVersionOnly = Value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the scope (first version only, last version only, or all versions)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property VersionScope() As VersionScope
      Get
        Return menuVersionScope
      End Get
      Set(ByVal value As VersionScope)
        menuVersionScope = value
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value determining whether or not the 
    ''' property should persist after the transformation is complete.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>If set to false then the property will be deleted once 
    ''' the transformation is complete.  This is used when a property 
    ''' is scoped only for the duration of the transformation, 
    ''' such as a temporary variable.</remarks>
    <XmlAttribute()>
    Public Property Persistent() As Boolean
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
      MyBase.New(ActionType.CreateProperty)
    End Sub

    Public Sub New(ByVal lpPropertyName As String, ByVal lpPropertyType As Core.PropertyType)
      MyBase.New(lpPropertyName, ActionType.CreateProperty)
      Try
        PropertyType = lpPropertyType
        Name = String.Format("Create{0}", lpPropertyName)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpSource As IProperty,
                   Optional ByVal lpFirstVersionOnly As Boolean = False,
                   Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty,
                   Optional ByVal lpPersistent As Boolean = True)
      Me.New(lpSource.Name, lpSource.Value, lpSource.Type, lpFirstVersionOnly, lpPropertyScope, lpPersistent)
    End Sub

    Public Sub New(ByVal lpPropertyName As String,
                   ByVal lpPropertyValue As Object,
                   ByVal lpPropertyType As Core.PropertyType,
                   Optional ByVal lpFirstVersionOnly As Boolean = False,
                   Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty,
                   Optional ByVal lpPersistent As Boolean = True)

      MyBase.New(lpPropertyName, ActionType.CreateProperty)
      Try
        Name = String.Format("Create{0}", lpPropertyName)
        PropertyScope = lpPropertyScope
        PropertyValue = lpPropertyValue
        PropertyType = lpPropertyType
        FirstVersionOnly = lpFirstVersionOnly
        Persistent = lpPersistent
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function Execute(ByRef lpErrorMessage As String) As ActionResult
      ' Create a new property
      'Dim lobjCreatePropertyAction As CreatePropertyAction = lobjAction

      Dim lobjActionResult As ActionResult = Nothing
      ' Try to create the property

      Try

        If TypeOf Transformation.Target Is Document Then
          'Check to see if property already exists, if it does, then do not create it again
          If (Transformation.Document.PropertyExists(PropertyScope, PropertyName, False)) Then
            'return
            Return New ActionResult(Me, False, String.Format("The property '{0}' already exists in document '{1}', skipping CreatePropertyAction.", PropertyName, Transformation.Document.Name))
          End If


          ' See if this is a script or not
          If Script.IsCtScript(PropertyValue) = True Then
            PropertyValue = Script.GetValue(PropertyValue, PropertyType)
          End If

          ' TODO: Add Document Property to base Action class
          If Transformation.Document.CreateProperty(PropertyName,
                                                 PropertyValue,
                                                 PropertyType,
                                                 Cardinality,
                                                 PropertyScope,
                                                 VersionScope,
                                                 Persistent) = True Then
            lobjActionResult = New ActionResult(Me, True)
          Else
            lobjActionResult = New ActionResult(Me, False, String.Format("Failed to create property '{0}' using value '{1}'", PropertyName, PropertyValue))
          End If


        ElseIf TypeOf Transformation.Target Is Folder Then
          'Check to see if property already exists, if it does, then do not create it again
          If (Transformation.Folder.PropertyExists(PropertyName, False)) Then
            'return
            Return New ActionResult(Me, False, String.Format("The property '{0}' already exists in folder '{1}', skipping CreatePropertyAction.", PropertyName, Transformation.Folder.Name))
          End If


          ' See if this is a script or not
          If Script.IsCtScript(PropertyValue) = True Then
            PropertyValue = Script.GetValue(PropertyValue, PropertyType)
          End If

          If Transformation.Folder.CreateProperty(PropertyName,
                                                 PropertyValue,
                                                 PropertyType,
                                                 Cardinality,
                                                 Persistent) = True Then
            lobjActionResult = New ActionResult(Me, True)
          Else
            lobjActionResult = New ActionResult(Me, False, String.Format("Failed to create property '{0}' using value '{1}'", PropertyName, PropertyValue))
          End If


        End If

        Return lobjActionResult

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
        Dim lobjParameters As IParameters = MyBase.GetDefaultParameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_CARDINALITY) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_CARDINALITY, Cardinality.ecmSingleValued, GetType(Cardinality),
            "Specifies the cardinality of the new property."))
        End If

        If lobjParameters.Contains(PARAM_FIRST_VERSION_ONLY) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_FIRST_VERSION_ONLY, False,
            "Specifies whether the new property should only be added to the first version of the document."))
        End If

        If lobjParameters.Contains(PARAM_PERSISTENT) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmBoolean, PARAM_PERSISTENT, True,
            "Specifies whether the new property should remain once the transformation has completed."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_TYPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PROPERTY_TYPE, PropertyType.ecmString, GetType(PropertyType),
            "Specifies the cardinality of the new property."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_VALUE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PROPERTY_VALUE, String.Empty,
            "Specifies the original value of the new property."))
        End If

        If lobjParameters.Contains(PARAM_VERSION_SCOPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_VERSION_SCOPE, VersionScope.AllVersions, GetType(VersionScope),
            "Specifies the version scope for the new property."))
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
        Me.FirstVersionOnly = GetBooleanParameterValue(PARAM_FIRST_VERSION_ONLY, False)
        Me.Persistent = GetBooleanParameterValue(PARAM_PERSISTENT, True)
        Me.PropertyType = GetEnumParameterValue(PARAM_PROPERTY_TYPE, GetType(PropertyType), PropertyType.ecmString)
        Me.PropertyValue = GetParameterValue(PARAM_PROPERTY_VALUE, String.Empty)
        Me.VersionScope = GetEnumParameterValue(PARAM_VERSION_SCOPE, GetType(VersionScope), VersionScope.AllVersions)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace