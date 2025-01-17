'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports System.Xml
Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class SingletonXmlProperty
    Inherits SingletonProperty

#Region "Public Properties"

    <XmlIgnore()>
    Public Overloads Property Value As String
      Get
        Try
          If MyBase.Value IsNot Nothing AndAlso TypeOf (MyBase.Value) Is XmlDocument Then
            Return DirectCast(MyBase.Value, XmlDocument).OuterXml
          ElseIf MyBase.Value IsNot Nothing Then
            Return MyBase.Value
          Else
            Return Nothing
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal lpValue As String)
        MyBase.Value = lpValue
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(PropertyType.ecmXml)
    End Sub

    Public Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmXml, lpName)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As XmlDocument)
      MyBase.New(PropertyType.ecmXml, lpName, lpSystemName, lpValue)
    End Sub

    Public Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpValue As String)
      MyBase.New(PropertyType.ecmXml, lpName, lpSystemName, lpValue)
    End Sub

#End Region

  End Class

End Namespace