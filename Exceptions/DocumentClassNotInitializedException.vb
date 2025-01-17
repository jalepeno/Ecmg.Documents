'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Providers

#End Region

Namespace Exceptions

  ''' <summary>
  ''' Exception defined for cases where a 
  ''' document class can't be deserialized from a RIF
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentClassNotInitializedException
    Inherits ItemNotInitializedException

#Region "Public Properties"

    Public Property DocumentClassName() As String
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

    Public Sub New(ByVal documentClassName As String,
                   ByVal contentSource As ContentSource)
      MyBase.New(String.Format("Document class '{0}' could not be initialized.", documentClassName), contentSource)
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal documentClassName As String,
                   ByVal contentSource As ContentSource)
      MyBase.New(message, documentClassName, contentSource)
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal documentClassName As String,
                   ByVal contentSource As ContentSource,
                   ByVal innerException As System.Exception)
      MyBase.New(message, documentClassName, contentSource, innerException)
    End Sub

#End Region

  End Class

End Namespace