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

  Public Class SingletonGuidProperty
    Inherits SingletonProperty

#Region "Public Properties"

    <XmlIgnore()>
    Public Overloads Property Value As Nullable(Of Guid)
      Get
        Try
          If MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is Guid Then
            Return MyBase.Value
          ElseIf MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is String Then
            If MyBase.Value.ToString.Length = 0 Then
              Return Nothing
            Else
              Dim lobjOutputGuid As Guid = Nothing
              If Helper.IsGuid(MyBase.Value, lobjOutputGuid) Then
                Return lobjOutputGuid
              Else
                Return Nothing
              End If
              'Return New Guid(MyBase.Value.ToString)
            End If
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Nullable(Of Guid))
        MyBase.Value = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmGuid)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmGuid, lpName)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As Nullable(Of Guid))
      MyBase.New(PropertyType.ecmGuid, lpName, lpSystemName, lpValue)
    End Sub

#End Region

  End Class

End Namespace