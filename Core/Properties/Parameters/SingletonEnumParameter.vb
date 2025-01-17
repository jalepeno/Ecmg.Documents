'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  SingletonEnumParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 3:42:27 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.ComponentModel
Imports Documents.TypeConverters
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class SingletonEnumParameter
    Inherits SingletonParameter

#Region "Class Variables"

    Private mobjEnumType As Type = Nothing

#End Region

#Region "Public Properties"

    <TypeConverter(GetType(EnumDescriptionTypeConverter))>
    Public Overloads Property Value As [Enum]
      Get
        Try
          Dim lenuParseTestValueResult As [Enum]
          If MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is [Enum] Then
            Return MyBase.Value
          ElseIf MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is String Then
            lenuParseTestValueResult = [Enum].Parse(EnumType, MyBase.Value)
            Return lenuParseTestValueResult
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As [Enum])
        MyBase.Value = value
      End Set
    End Property

    Public ReadOnly Property PossibleValues As IEnumerable
      Get
        Try
          Return GetPossibleValues()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Private Function GetPossibleValues() As IEnumerable
      Try
        Dim lenuValues As IEnumerable
        If EnumType IsNot Nothing Then
          lenuValues = [Enum].GetValues(EnumType)
        Else
          lenuValues = Nothing
        End If
        Return lenuValues
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

    Public Property EnumValue As [Enum]
      Get
        Try
          Return [Enum].Parse(Me.EnumType, Me.Value.ToString())
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As [Enum])
        Try
          MyBase.Value = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public Shadows ReadOnly Property EnumType As Type
      Get
        Try
          If mobjEnumType Is Nothing OrElse mobjEnumType.Name = "Int32" Then
            mobjEnumType = GetEnumType()
          End If
          Return mobjEnumType
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public ReadOnly Property EnumName As String
      Get
        Try
          Return mstrEnumTypeName
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Sub SetStandardValues(lpStandardValues As IEnumerable)
      Try
        mobjStandardValues = lpStandardValues

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmEnum)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmEnum, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As [Enum])
      MyBase.New(PropertyType.ecmEnum, lpName, lpSystemName, lpValue)
      Try
        mobjEnumType = lpValue.GetType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    'Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String,
    '               ByVal lpValue As String, ByVal lpEnumType As Type)
    '  MyBase.New(PropertyType.ecmEnum, lpName)
    '  Try
    '    mobjEnumType = lpEnumType
    '    Me.SystemName = lpSystemName
    '    Me.Value = [Enum].Parse(lpEnumType, lpValue)
    '    If lpEnumType IsNot Nothing Then
    '      mobjStandardValues = Helper.EnumerationDictionary(lpEnumType).Keys
    '    End If
    '  Catch ex As Exception
    '    ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '    ' Re-throw the exception to the caller
    '    Throw
    '  End Try
    'End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String,
                   ByVal lpValue As [Enum], ByVal lpEnumType As Type)
      MyBase.New(PropertyType.ecmEnum, lpName, lpSystemName, lpValue)
      Try
        mobjEnumType = lpEnumType
        If lpEnumType IsNot Nothing Then
          mstrEnumTypeName = mobjEnumType.Name
          mobjStandardValues = Helper.EnumerationDictionary(lpEnumType).Keys
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Private Methods"

    Friend Sub SetEnumType(lpType As Type)
      Try
        mobjEnumType = lpType
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Private Function GetEnumType() As Type
      Try
        If String.IsNullOrEmpty(EnumTypeName) Then
          Return Nothing
        Else
          'Return System.Type.GetType(EnumTypeName)
          Return Helper.GetTypeFromAssembly(Reflection.Assembly.GetExecutingAssembly, EnumTypeName)
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