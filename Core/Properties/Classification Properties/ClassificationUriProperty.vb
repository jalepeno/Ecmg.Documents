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

  Public Class ClassificationUriProperty
    Inherits ClassificationProperty

#Region "Public Properties"

    <XmlIgnore()>
    Public Overloads Property DefaultValue As Uri

    <XmlElement("DefaultValue")>
    Public Property DefaultValueString As String
      Get
        Return DefaultValue.ToString
      End Get
      Set(ByVal value As String)
        DefaultValue = New Uri(value)
      End Set
    End Property

#End Region

#Region "Constructors"

    Friend Sub New()
      MyBase.New(PropertyType.ecmUri)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmUri, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpDefaultValue As Uri)

      MyBase.New(PropertyType.ecmUri, lpName, lpSystemName)

      Try

        If lpDefaultValue IsNot Nothing Then
          DefaultValue = lpDefaultValue
        Else
          DefaultValue = Nothing
        End If

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

    Friend Sub New(ByVal lpName As String,
                   ByVal lpSystemName As String,
                   ByVal lpDefaultValue As String)
      MyBase.New(PropertyType.ecmUri, lpName, lpSystemName)
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

