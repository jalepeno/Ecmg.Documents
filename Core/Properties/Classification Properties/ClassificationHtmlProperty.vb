﻿'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ClassificationHtmlProperty
    Inherits ClassificationProperty

#Region "Public Properties"

    Public Overloads Property DefaultValue As String

#End Region

#Region "Constructors"

    Friend Sub New()
      MyBase.New(PropertyType.ecmHtml)
    End Sub

    Friend Sub New(ByVal lpName As String)
      MyBase.New(PropertyType.ecmHtml, lpName)
    End Sub

    Friend Sub New(ByVal lpName As String, ByVal lpSystemName As String, ByVal lpDefaultValue As String)

      MyBase.New(PropertyType.ecmHtml, lpName, lpSystemName)

      Try

        If lpDefaultValue IsNot Nothing Then
          DefaultValue = lpDefaultValue
        Else
          DefaultValue = String.Empty
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
