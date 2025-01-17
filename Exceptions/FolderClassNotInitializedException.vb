' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FolderClassNotInitializedException.vb
'  Description :  [type_description_here]
'  Created     :  3/5/2012 1:21:21 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Providers

#End Region

Namespace Exceptions

  Public Class FolderClassNotInitializedException
    Inherits ItemNotInitializedException

#Region "Public Properties"

    Public Property FolderClassName() As String
      Get
        Return MyBase.ItemName
      End Get
      Set(ByVal value As String)
        MyBase.ItemName = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal folderClassName As String,
                   ByVal contentSource As ContentSource)
      MyBase.New(String.Format("Folder class '{0}' could not be initialized.", folderClassName), contentSource)
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal folderClassName As String,
                   ByVal contentSource As ContentSource)
      MyBase.New(message, folderClassName, contentSource)
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal folderClassName As String,
                   ByVal contentSource As ContentSource,
                   ByVal innerException As System.Exception)
      MyBase.New(message, folderClassName, contentSource, innerException)
    End Sub

#End Region

  End Class

End Namespace