' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentStreamContainer.vb
'  Description :  [type_description_here]
'  Created     :  12/14/2011 8:18:00 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports System.IO
Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ContentStreamContainer
    Inherits ContentContainerBase

#Region "Public Properties"

    Public Overrides ReadOnly Property CanRead As Boolean
      Get
        Try
          If ContentStream IsNot Nothing AndAlso ContentStream.CanRead Then
            Return True
          Else
            Return False
          End If
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
    End Property

    Public Property ContentStream As Stream
      Get
        Return MyBase.FileContent
      End Get
      Set(ByVal value As Stream)
        MyBase.FileContent = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal lpFileName As String, ByVal lpMimeType As String, ByVal lpContentStream As Stream)
      MyBase.New(lpFileName, lpMimeType, lpContentStream)
    End Sub

    Public Sub New(ByVal lpContent As Content)
      MyBase.New(lpContent)
    End Sub

#End Region

  End Class

End Namespace

