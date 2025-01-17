'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  SingletonBooleanParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 2:03:15 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class SingletonBooleanParameter
    Inherits SingletonParameter

#Region "Public Properties"

    '<XmlIgnore()> _
    'Public Overloads Property Value As Nullable(Of Boolean)
    '  Get
    '    Try
    '      If MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is Boolean Then
    '        Return MyBase.Value
    '      ElseIf MyBase.Value IsNot Nothing Then
    '        If TypeOf (MyBase.Value) Is String AndAlso MyBase.Value.ToString.Length = 0 Then
    '          Return Nothing
    '        Else
    '          Dim lblnValue As Boolean
    '          If Boolean.TryParse(MyBase.Value, lblnValue) = True Then
    '            Return lblnValue
    '          Else
    '            Return Nothing
    '          End If
    '        End If
    '        Return CBool(MyBase.Value)
    '      Else
    '        Return Nothing
    '      End If
    '    Catch ex As Exception
    '      ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
    '      ' Re-throw the exception to the caller
    '      Throw
    '    End Try
    '  End Get
    '  Set(ByVal value As Nullable(Of Boolean))
    '    If value.ToString.Length = 0 Then
    '      MyBase.Value = Nothing
    '    Else
    '      Dim lblnValue As Boolean
    '      If Boolean.TryParse(value, lblnValue) = True Then
    '        MyBase.Value = lblnValue
    '      Else
    '        MyBase.Value = Nothing
    '      End If
    '    End If
    '  End Set
    'End Property

    Public Overloads Property Value As Boolean
      Get
        Try
          If MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is Boolean Then
            Return MyBase.Value
          ElseIf MyBase.Value IsNot Nothing Then
            If TypeOf (MyBase.Value) Is String AndAlso MyBase.Value.ToString.Length = 0 Then
              Return False
            Else
              Dim lblnValue As Boolean
              If Boolean.TryParse(MyBase.Value, lblnValue) = True Then
                Return lblnValue
              Else
                Return False
              End If
            End If
            Return CBool(MyBase.Value)
          Else
            Return False
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Boolean)
        Try
          If value.ToString.Length = 0 Then
            MyBase.Value = False
          Else
            Dim lblnValue As Boolean
            If Boolean.TryParse(value, lblnValue) = True Then
              MyBase.Value = lblnValue
            Else
              MyBase.Value = False
            End If
          End If
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
      MyBase.New(PropertyType.ecmBoolean)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmBoolean, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String)
      MyBase.New(PropertyType.ecmBoolean, lpName, lpSystemName, Nothing)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As Nullable(Of Boolean))
      MyBase.New(PropertyType.ecmBoolean, lpName, lpSystemName, Nothing)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As String)
      MyBase.New(PropertyType.ecmBoolean, lpName, lpSystemName, Nothing)
      Try
        Dim lblnValue As Boolean
        If lpValue IsNot Nothing Then
          If Boolean.TryParse(lpValue, lblnValue) = True Then
            Value = lblnValue
          Else
            Throw New InvalidCastException(String.Format("Unable to cast value '{0}' as boolean for parameter '{1}'", lpValue, lpName))
          End If
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

#Region "Overriden Methods"

    Public Overrides ReadOnly Property HasStandardValues() As Boolean
      Get
        Return True
      End Get
    End Property

    Public Overrides ReadOnly Property StandardValues() As IEnumerable
      Get
        Try
          Return SingletonBooleanProperty.GetStandardValues()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Protected Overrides Function PropertyHasValue() As Boolean
      Try
        Return True
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace