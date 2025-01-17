'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Utilities

Namespace Core
  ''' <exclude/>
  <Microsoft.VisualBasic.HideModuleName()>
  <Serializable()>
  Public Class HeaderProperty
    Inherits ECMProperty

#Region "Class Enumerations"

    ''' <summary>Determines the stage at which the property may be set</summary>
    Public Enum PropertyMutability
      ''' <summary>The property may be set at any time</summary>
      ReadWrite = 0
      ''' <summary>The property may be set only at the time of creation</summary>
      SettableOnlyOnCreate = 1
      ''' <summary>The property may be set only at the time of transformation</summary>
      SettableOnlyOnTransform = 2
    End Enum

#End Region

#Region "Class Variables"

    Private mblnMutability As PropertyMutability = PropertyMutability.SettableOnlyOnCreate

#End Region

#Region "Public Properties"

    Public Property Mutability() As PropertyMutability
      Get
        Return mblnMutability
      End Get
      Set(ByVal value As PropertyMutability)
        Try
          If Helper.CallStackContainsMethodName("New", "Deserialize") Then
            mblnMutability = value
          Else
            Throw New InvalidOperationException("Although 'Mutability' is a public property, set operations are not allowed.  Treat property as read-only.")
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Overrides Methods"

    Public Overloads Property Value() As Object
      Get
        Return MyBase.Value
      End Get
      Set(ByVal value As Object)
        Try
          Select Case Mutability
            Case PropertyMutability.ReadWrite
              MyBase.Value = value
            Case PropertyMutability.SettableOnlyOnCreate
              If Helper.CallStackContainsMethodName("Deserialize") Then
                MyBase.Value = value
              Else
                Throw New InvalidOperationException("Although 'Value' is a public property, the Mutability is 'SettableOnlyOnCreate'.  Therefore set operations are not allowed.  Treat property as read-only.")
              End If
            Case PropertyMutability.SettableOnlyOnTransform
              If Helper.CallStackContainsMethodName("Deserialize", "Transform") Then
                MyBase.Value = value
              Else
                Throw New InvalidOperationException("Although 'Value' is a public property, the Mutability is 'SettableOnlyOnTransform'.  Therefore set operations are not allowed.  Treat property as read-only.")
              End If
          End Select
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          '  Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(Nothing)
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpCardinality As Cardinality)

      MyBase.New(Nothing)

      Try
        Type = lpMenuType
        Name = lpName
        SystemName = lpName
        Cardinality = lpCardinality
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try

    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValue As Object)

      MyBase.New(Nothing)

      Try
        Type = lpMenuType
        Name = lpName
        Value = lpValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValues As Values)

      MyBase.New(Nothing)

      Try
        Type = lpMenuType
        Name = lpName
        Values = lpValues
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpCardinality As Cardinality,
                   ByVal lpDefaultValue As Object)

      MyBase.New(Nothing)

      Try
        Type = lpMenuType
        Name = lpName
        Cardinality = lpCardinality
        DefaultValue = lpDefaultValue
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValue As Object,
                   ByVal lpDefaultValue As Object)

      MyBase.New(Nothing)

      Try

        Type = lpMenuType
        Name = lpName
        Value = lpValue
        DefaultValue = lpDefaultValue

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    ' With Mutability
    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpCardinality As Cardinality,
                   ByVal lpMutability As PropertyMutability)

      MyBase.New(Nothing)

      Try

        Type = lpMenuType
        Name = lpName
        Cardinality = lpCardinality
        mblnMutability = lpMutability

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValue As Object,
                   ByVal lpMutability As PropertyMutability)

      MyBase.New(Nothing)

      Try

        Type = lpMenuType
        Name = lpName
        Value = lpValue
        mblnMutability = lpMutability

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValues As Values,
                   ByVal lpMutability As PropertyMutability)

      MyBase.New(Nothing)

      Try

        Type = lpMenuType
        Name = lpName
        Values = lpValues
        mblnMutability = lpMutability

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpCardinality As Cardinality,
                   ByVal lpDefaultValue As Object,
                   ByVal lpMutability As PropertyMutability)

      MyBase.New(Nothing)

      Try

        Type = lpMenuType
        Name = lpName
        Cardinality = lpCardinality
        DefaultValue = lpDefaultValue
        mblnMutability = lpMutability

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Public Sub New(ByVal lpMenuType As PropertyType,
                   ByVal lpName As String,
                   ByVal lpValue As Object,
                   ByVal lpDefaultValue As Object,
                   ByVal lpMutability As PropertyMutability)

      MyBase.New(Nothing)

      Try

        Type = lpMenuType
        Name = lpName
        Value = lpValue
        DefaultValue = lpDefaultValue
        mblnMutability = lpMutability

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        '  Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace