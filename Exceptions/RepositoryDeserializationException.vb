'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Exceptions

  Public Class RepositoryDeserializationException
    Inherits CtsException

#Region "Class Variables"

    Private mstrRepositoryFilePath As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property RepositoryFilePath() As String
      Get
        Return mstrRepositoryFilePath
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal repositoryFilePath As String)
      MyBase.New()
      mstrRepositoryFilePath = repositoryFilePath
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal repositoryFilePath As String)
      MyBase.New(message)
      mstrRepositoryFilePath = repositoryFilePath
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal repositoryFilePath As String,
                   ByVal innerexception As Exception)
      MyBase.New(message, innerexception)
      mstrRepositoryFilePath = repositoryFilePath
    End Sub

#End Region

  End Class

End Namespace
