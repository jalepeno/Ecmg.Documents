'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  SingletonLongParameter.vb
'   Description :  [type_description_here]
'   Created     :  6/18/2013 2:11:40 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class SingletonLongParameter
    Inherits SingletonParameter

#Region "Public Properties"

    <XmlIgnore()>
    Public Overloads Property Value As Nullable(Of Long)
      Get
        Try
          Dim llngParseTestValueResult As Long
          If MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is Long Then
            Return MyBase.Value
          ElseIf MyBase.Value IsNot Nothing AndAlso Long.TryParse(MyBase.Value, llngParseTestValueResult) Then
            Return Long.Parse(MyBase.Value)
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As Nullable(Of Long))
        MyBase.Value = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmLong)
    End Sub

    Protected Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmLong, lpName)
    End Sub

    Protected Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As Nullable(Of Long))
      MyBase.New(PropertyType.ecmLong, lpName, lpSystemName, lpValue)
    End Sub

#End Region

  End Class

End Namespace