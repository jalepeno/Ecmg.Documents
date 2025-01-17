' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ContentBytesContainer.vb
'  Description :  [type_description_here]
'  Created     :  12/14/2011 8:27:00 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Utilities

#End Region

Namespace Core

  Public Class ContentBytesContainer
    Inherits ContentContainerBase

#Region "Public Properties"

    Public Overrides ReadOnly Property CanRead As Boolean
      Get
        Try
          If ContentBytes IsNot Nothing AndAlso ContentBytes.Length > 0 Then
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

    Public Property ContentBytes As Byte()
      Get
        Return MyBase.FileContent
      End Get
      Set(ByVal value As Byte())
        MyBase.FileContent = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpFileName As String, ByVal lpMimeType As String, ByVal lpContentStream As Byte())
      MyBase.New(lpFileName, lpMimeType, lpContentStream)
    End Sub

    Public Sub New(ByVal lpContent As Content)
      Try
        FileName = lpContent.FileName
        MimeType = lpContent.MIMEType
        If lpContent.CanRead Then
          ContentBytes = lpContent.ToByteArray
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
