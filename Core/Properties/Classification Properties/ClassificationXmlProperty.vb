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

  Public Class ClassificationXmlProperty
    Inherits ClassificationProperty

#Region "Public Properties"

    'Public Overloads Property DefaultValue As String

    <XmlIgnore()>
    Public Overloads Property DefaultValue As XmlDocument

    <XmlElement("DefaultValue")>
    Public Property DefaultValueString As String
      Get
        Try
          If DefaultValue IsNot Nothing Then
            Return DefaultValue.ToString
          Else
            Return String.Empty
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(ByVal value As String)
        Try
          DefaultValue = New XmlDocument
          If value IsNot Nothing Then
            DefaultValue.LoadXml(value)
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

    Friend Sub New()
      MyBase.New(PropertyType.ecmXml)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmXml, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpDefaultValue As String)

      MyBase.New(PropertyType.ecmXml, lpName, lpSystemName)

      Try

        If lpDefaultValue IsNot Nothing Then
          DefaultValueString = lpDefaultValue
        Else
          DefaultValue = Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace
