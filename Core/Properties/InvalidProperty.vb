'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  <DebuggerDisplay("{DebuggerIdentifier(),nq}")>
  Public Class InvalidProperty
    Implements IProperty
    Implements IComparable

#Region "Class Enumerations"

    Public Enum InvalidPropertyScope
      Document = 0
      FirstVersion = 1
      AllVersions = 2
      AllExceptFirstVersion = 3
    End Enum

#End Region

#Region "Class Variables"

    Private mobjBaseProperty As ECMProperty
    Private menuScope As InvalidPropertyScope

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Gets the PackedName of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns>Returns BaseProperty.PackedName if BaseProperty 
    ''' is not null, otherwise returns String.Empty.</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property PackedName As String
      Get
        If BaseProperty IsNot Nothing Then
          Return BaseProperty.PackedName
        Else
          Return String.Empty
        End If
      End Get
    End Property

    Public ReadOnly Property BaseProperty As ECMProperty
      Get
        Return mobjBaseProperty
      End Get
    End Property

    Public ReadOnly Property Scope As InvalidPropertyScope
      Get
        Return menuScope
      End Get
    End Property

    Public Overridable ReadOnly Property HasStandardValues As Boolean Implements IProperty.HasStandardValues
      Get
        Try
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
          Return New List(Of String)
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property Value As Object Implements IProperty.Value
      Get
        Return BaseProperty.Value
      End Get
      Set(value As Object)
        BaseProperty.Value = value
      End Set
    End Property

    Public Property Values As Object Implements IProperty.Values
      Get
        Return BaseProperty.Values
      End Get
      Set(value As Object)
        BaseProperty.Values = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpBaseProperty As ECMProperty, ByVal lpScope As InvalidPropertyScope)
      Try
        mobjBaseProperty = lpBaseProperty
        menuScope = lpScope
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Public Methods"

    Public Overrides Function ToString() As String Implements IProperty.ToDebugString
      Try
        Return DebuggerIdentifier()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Private Methods"

    Protected Overridable Function DebuggerIdentifier() As String
      Try
        If BaseProperty IsNot Nothing Then
          Return String.Format("{0} ({1})", Me.BaseProperty.ToString, Me.Scope.ToString)
        Else
          Return Me.GetType.Name
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IComparable Implementation"

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
      Try
        If TypeOf obj Is IProperty Then
          Return Name.CompareTo(obj.Name)
        Else
          Throw New ArgumentException("Object is not an IProperty")
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "IProperty Implementation"

    ''' <summary>
    ''' Clears current value(s) from the property.
    ''' </summary>
    ''' <remarks>
    ''' Clears both the Value and Values properties
    ''' </remarks>
    Public Sub Clear() Implements IProperty.Clear
      Try
        Me.BaseProperty.Clear()
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ''' <summary>
    ''' Gets or sets the Cardinality of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Cardinality As Cardinality Implements IProperty.Cardinality
      Get
        Try
          Return BaseProperty.Cardinality
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Cardinality)
        Try
          BaseProperty.Cardinality = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the DefaultValue of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DefaultValue As Object Implements IProperty.DefaultValue
      Get
        Try
          Return BaseProperty.DefaultValue
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Object)
        Try
          BaseProperty.DefaultValue = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the Description of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Description As String Implements IProperty.Description
      Get
        Try
          Return BaseProperty.Description
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          BaseProperty.Description = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the DisplayName of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DisplayName() As String Implements IProperty.DisplayName
      Get
        Try
          Return BaseProperty.DisplayName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As String)
        Try
          BaseProperty.DisplayName = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the HasValue property of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property HasValue As Boolean Implements IProperty.HasValue
      Get
        Try
          Return BaseProperty.HasValue
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the Name of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Name As String Implements IProperty.Name
      Get
        Try
          Return BaseProperty.Name
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          BaseProperty.Name = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets the Persistent property of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Persistent As Boolean Implements IProperty.Persistent
      Get
        Try
          Return BaseProperty.Persistent
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    ''' <summary>
    ''' Gets or sets the SystemName of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SystemName As String Implements IProperty.SystemName
      Get
        Try
          Return BaseProperty.SystemName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          BaseProperty.SystemName = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    ''' <summary>
    ''' Gets or sets the Type of the BaseProperty
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Type As PropertyType Implements IProperty.Type
      Get
        Try
          Return BaseProperty.Type
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As PropertyType)
        Try

        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

  End Class

End Namespace