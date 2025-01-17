'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Xml.Serialization
Imports Documents.Core
Imports Documents.Utilities

Namespace Transformations
  ''' <summary>Base class for actions which modify properties of documents.</summary>
  <XmlInclude(GetType(ChangePropertyValue)),
    XmlInclude(GetType(CreatePropertyAction)),
    XmlInclude(GetType(DeletePropertyAction)),
    XmlInclude(GetType(RenamePropertyAction)),
    XmlInclude(GetType(SplitPropertyAction)),
    XmlInclude(GetType(DeletePropertiesNotInDocumentClass)),
    XmlInclude(GetType(ChangePropertyCardinality)),
    XmlInclude(GetType(ChangeAllTimesToUTC)),
    XmlInclude(GetType(RemoveTimeFromAllDatesAction)),
    XmlInclude(GetType(ClearAllPropertyValues))>
  <Serializable()>
  Public MustInherit Class PropertyAction
    Inherits Action

#Region "Class Constants"

    Friend Const PARAM_PROPERTY_NAME As String = "PropertyName"
    Friend Const PARAM_PROPERTY_SCOPE As String = "PropertyScope"

#End Region

#Region "Class Enumerations"

    Public Enum RepositoryScope

      ''' <summary>
      ''' References the source repository.
      ''' </summary>
      ''' <remarks></remarks>
      Source

      ''' <summary>
      ''' References the destination repository.
      ''' </summary>
      ''' <remarks></remarks>
      Destination

    End Enum

#End Region

#Region "Class Variables"

    Private mstrPropertyName As String = String.Empty
    Private menuPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty

#End Region

#Region "Public Properties"

    <XmlAttribute()>
    Public Overrides Property Name() As String
      Get
        Try

          If mstrName.Length = 0 Then
            If PropertyName IsNot Nothing AndAlso PropertyName.Length > 0 Then
              mstrName = String.Format("{0}({1})", Me.GetType.Name, PropertyName)
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

    Public Property PropertyName() As String
      Get
        Return mstrPropertyName
      End Get
      Set(ByVal Value As String)
        mstrPropertyName = Value
      End Set
    End Property

    Public Property PropertyScope() As Core.PropertyScope
      Get
        Return menuPropertyScope
      End Get
      Set(ByVal Value As Core.PropertyScope)
        menuPropertyScope = Value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpActionType As ActionType, Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty)
      MyBase.New(lpActionType, "")
      Try
        PropertyScope = lpPropertyScope
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpPropertyName As String, Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty)
      Try
        PropertyName = lpPropertyName
        PropertyScope = lpPropertyScope
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpPropertyName As String, ByVal lpActionType As ActionType, Optional ByVal lpPropertyScope As Core.PropertyScope = Core.PropertyScope.VersionProperty)
      MyBase.New(lpActionType, lpPropertyName)
      Try
        PropertyName = lpPropertyName
        PropertyScope = lpPropertyScope
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Protected Methods"

    Protected Overrides Function GetDefaultParameters() As IParameters
      Try
        Dim lobjParameters As IParameters = New Parameters
        Dim lobjParameter As IParameter = Nothing

        If lobjParameters.Contains(PARAM_PROPERTY_NAME) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmString, PARAM_PROPERTY_NAME, PropertyName,
            "Specifies the name of the target property to modify."))
        End If

        If lobjParameters.Contains(PARAM_PROPERTY_SCOPE) = False Then
          lobjParameters.Add(ParameterFactory.Create(PropertyType.ecmEnum, PARAM_PROPERTY_SCOPE, PropertyScope, GetType(PropertyScope),
            "Specifies the scope of the target property."))
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
        Me.PropertyName = GetStringParameterValue(PARAM_PROPERTY_NAME, String.Empty)
        Me.PropertyScope = GetEnumParameterValue(PARAM_PROPERTY_SCOPE, GetType(PropertyScope), PropertyScope.VersionProperty)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Overridable Function GetFolderMetaHolder() As IMetaHolder
      Try
        Return Transformation.Folder
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Protected Overridable Function GetMetaHolder() As IMetaHolder
      Try
        Return GetMetaHolder(PropertyScope)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller.
        Throw
      End Try
    End Function

    Protected Overridable Function GetMetaHolder(ByVal lpPropertyScope As PropertyScope) As IMetaHolder
      Try
        Dim lobjMetaHolder As IMetaHolder = Nothing

        Select Case lpPropertyScope
          Case Core.PropertyScope.DocumentProperty
            lobjMetaHolder = Transformation.Document
          Case Core.PropertyScope.VersionProperty
            lobjMetaHolder = Transformation.Document.GetFirstVersion
          Case Core.PropertyScope.ContentProperty
            lobjMetaHolder = Transformation.Document.GetFirstVersion.Contents(0)
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

  End Class

End Namespace