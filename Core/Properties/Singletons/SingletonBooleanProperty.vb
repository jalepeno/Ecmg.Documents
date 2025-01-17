'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class SingletonBooleanProperty
    Inherits SingletonProperty

#Region "Public Properties"

    <XmlIgnore()>
    Public Overloads Property Value As Nullable(Of Boolean)
      Get
        Try
          If MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is Boolean Then
            Return MyBase.Value
          ElseIf MyBase.Value IsNot Nothing Then
            If TypeOf (MyBase.Value) Is String AndAlso MyBase.Value.ToString.Length = 0 Then
              Return Nothing
            Else
              Dim lblnValue As Boolean
              If Boolean.TryParse(MyBase.Value, lblnValue) = True Then
                Return lblnValue
              Else
                Return Nothing
              End If
            End If
            Return CBool(MyBase.Value)
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Nullable(Of Boolean))
        If value.ToString.Length = 0 Then
          MyBase.Value = Nothing
        Else
          Dim lblnValue As Boolean
          If Boolean.TryParse(value, lblnValue) = True Then
            MyBase.Value = lblnValue
          Else
            MyBase.Value = Nothing
          End If
        End If
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmBoolean)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmBoolean, lpName)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpSystemName As String)
      MyBase.New(PropertyType.ecmBoolean, lpName, lpSystemName, Nothing)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As Nullable(Of Boolean))
      MyBase.New(PropertyType.ecmBoolean, lpName, lpSystemName, Nothing)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As String)
      MyBase.New(PropertyType.ecmBoolean, lpName, lpSystemName, Nothing)
      Try
        Dim lblnValue As Boolean
        If lpValue IsNot Nothing Then
          If Boolean.TryParse(lpValue, lblnValue) = True Then
            Value = lblnValue
          Else
            Throw New InvalidCastException(String.Format("Unable to cast value '{0}' as boolean for property '{1}'", lpValue, lpName))
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

#If CTSDOTNET = 1 Then

        Public Overrides ReadOnly Property StandardValues() As IEnumerable
      Get
        Try
          Return GetStandardValues()
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

#End If

    Protected Overrides Function PropertyHasValue() As Boolean
      Try
        If Value.HasValue Then
          Return True
        Else
          Return False
        End If
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

#Region "Public Shared Methods"

    Public Shared Function GetStandardValues() As IEnumerable
      Try
        Dim lobjStandardValues As New List(Of Boolean)
        lobjStandardValues.Add(False)
        lobjStandardValues.Add(True)
        Return lobjStandardValues
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Function

#End Region

  End Class

End Namespace