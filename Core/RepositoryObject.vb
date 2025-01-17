'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Runtime.Serialization
Imports System.Text
Imports System.Xml.Serialization
Imports Documents.Security
Imports Documents.Utilities

#End Region

Namespace Core

  ''' <summary>
  ''' Base class from which many Cts repository object types inherit
  ''' </summary>
  ''' <remarks></remarks>
  <DataContract()>
  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public MustInherit Class RepositoryObject
    Implements IComparable
    Implements IRepositoryObject

#Region "Class Variables"

    Private mstrID As String = String.Empty
    Private mstrName As String = String.Empty
    Private mstrDisplayName As String = String.Empty
    Private mstrDescriptiveText As String = String.Empty
    Protected WithEvents mobjProperties As New ECMProperties

    <DebuggerBrowsable(DebuggerBrowsableState.Never)>
    Private mobjPermissions As IPermissions = New Permissions

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The ID of the object.
    ''' </summary>
    <XmlAttribute()>
    <DataMember()>
    Public Property Id() As String Implements IRepositoryObject.Id
      Get
        Return mstrID
      End Get
      Set(ByVal value As String)
        mstrID = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
      End Set
    End Property

    ''' <summary>
    ''' The name of the object.
    ''' </summary>
    <DataMember()>
    Public Overridable Property Name() As String Implements IRepositoryObject.Name
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
      End Set
    End Property

    ''' <summary>
    ''' The description of the object.
    ''' </summary>
    <DataMember()>
    Public Property DescriptiveText() As String Implements IRepositoryObject.DescriptiveText
      Get
        Return mstrDescriptiveText
      End Get
      Set(ByVal value As String)
        mstrDescriptiveText = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
      End Set
    End Property

    ''' <summary>
    ''' The display name of the object.
    ''' </summary>
    ''' <returns>
    ''' The display name if available, otherwise 
    ''' a space trimmed copy of the name.
    ''' </returns>
    <DataMember()>
    Public Property DisplayName() As String Implements IRepositoryObject.DisplayName
      Get
        If Not String.IsNullOrEmpty(mstrDisplayName) Then
          Return mstrDisplayName
        ElseIf Not String.IsNullOrEmpty(Name) Then
          Return Trim(Name)
        Else
          Return String.Empty
        End If
      End Get
      Set(ByVal value As String)
        mstrDisplayName = value
        AddProperty(Reflection.MethodBase.GetCurrentMethod, value)
      End Set
    End Property

    '<DataMember()> _
    Public Property Properties() As ECMProperties
      Get
        Return mobjProperties
      End Get
      Set(ByVal value As ECMProperties)
        mobjProperties = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the permissions for the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Permissions As IPermissions Implements IRepositoryObject.Permissions
      Get
        Return mobjPermissions
      End Get
    End Property

    ''' <summary>
    ''' Gets the set of properties for the object
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property IProperties As IProperties Implements IRepositoryObject.Properties
      Get
        Return Properties
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Name As String)
      ' We will set the Name and the DisplayName to the same value to start.
      Try
        MyClass.Name = Name
        MyClass.DisplayName = Name
      Catch ex As Exception
        ApplicationLogging.LogException(ex, String.Format("RepositoryObject::New('{0}')", Name))
      End Try
    End Sub

#End Region

#Region "Protected Methods"

#Region "AddProperty"

    Protected Overridable Sub AddProperty(ByVal lpProperty As Reflection.MethodBase, ByVal lpPropertyValue As Object)
      Try

        Dim lstrPropertyName As String = lpProperty.Name

        If (lstrPropertyName.ToLower.StartsWith("set_")) OrElse (lstrPropertyName.ToLower.StartsWith("get_")) Then
          lstrPropertyName = lstrPropertyName.Substring(4)
        End If

        AddProperty(lstrPropertyName, lpPropertyValue)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Overridable Sub AddProperty(ByVal lpProperty As Reflection.MethodBase, ByVal lpPropertyValue As Object, ByVal lpPropertyType As Core.PropertyType)
      Try

        Dim lstrPropertyName As String = lpProperty.Name

        If (lstrPropertyName.ToLower.StartsWith("set_")) OrElse (lstrPropertyName.ToLower.StartsWith("get_")) Then
          lstrPropertyName = lstrPropertyName.Substring(4)
        End If

        AddProperty(lstrPropertyName, lpPropertyValue, lpPropertyType)

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Protected Overridable Sub AddProperty(ByVal lpPropertyName As String, ByVal lpPropertyValue As Object)
      Try
        AddProperty(lpPropertyName, lpPropertyValue, PropertyType.ecmString)
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Adds a property to the object 
    ''' </summary>
    ''' <param name="lpPropertyName">The name of the property</param>
    ''' <param name="lpPropertyValue">The property value</param>
    ''' <param name="lpPropertyType">The data type of the property</param>
    ''' <remarks>If a property with the specified name is already in the collection of properties, it wil not be added.  
    ''' The existing property will be retained and the current value will be updated with the specified value.</remarks>
    Protected Overridable Sub AddProperty(ByVal lpPropertyName As String, ByVal lpPropertyValue As Object, ByVal lpPropertyType As Core.PropertyType)
      Try

        If (mobjProperties Is Nothing) Then
          mobjProperties = New ECMProperties()
        End If

        If Properties.PropertyExists(lpPropertyName) = False Then
          'Dim lobjProperty As New ECMProperty(lpPropertyType, lpPropertyName, Cardinality.ecmSingleValued)
          'lobjProperty.Value = lpPropertyValue
          Dim lobjProperty As ECMProperty = PropertyFactory.Create(lpPropertyType, lpPropertyName, lpPropertyName, Cardinality.ecmSingleValued, lpPropertyValue)

          Properties.Add(lobjProperty)
        Else
          Properties(lpPropertyName).Value = lpPropertyValue
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

    Protected Friend Overridable Function DebuggerIdentifier() As String
      Dim lobjIdentifierBuilder As New StringBuilder
      Try
        Dim lstrName As String = Name

        ' 1st show the class, if assigned.
        If DisplayName.Length > 0 Then
          lobjIdentifierBuilder.AppendFormat("DisplayName: {0} ", DisplayName)
        End If

        If String.IsNullOrEmpty(Name) Then
          lobjIdentifierBuilder.Append("; Name not set")
        Else
          lobjIdentifierBuilder.AppendFormat("; Name: {0}", Name)
        End If

        If Id Is Nothing OrElse Id.Length = 0 Then
          lobjIdentifierBuilder.Append("; Id not set")
        Else
          lobjIdentifierBuilder.AppendFormat("; Id: {0}", Id)
        End If

        Return lobjIdentifierBuilder.ToString.TrimStart("; ")

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        Return lobjIdentifierBuilder.ToString
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
      Try
        If TypeOf obj Is RepositoryObject Then
          Return Name.CompareTo(obj.Name)
        Else
          Throw New ArgumentException("Object is not a RepositoryObject")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace