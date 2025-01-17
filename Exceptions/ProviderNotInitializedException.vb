'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Providers

Namespace Exceptions
  ''' <summary>
  ''' Exception defined for cases where there is not a valid provider reference
  ''' </summary>
  ''' <remarks></remarks>
  Public Class ProviderNotInitializedException
    Inherits CtsException

#Region "Class Variables"

    Private m_contentSource As ContentSource

#End Region

#Region "Public Properties"

    Public Property ContentSource() As ContentSource
      Get
        Return m_contentSource
      End Get
      Set(ByVal value As ContentSource)
        m_contentSource = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
    End Sub

    Public Sub New(ByVal contentSource As ContentSource)
      Me.ContentSource = contentSource
    End Sub

    Public Sub New(ByVal message As String, ByVal contentSource As ContentSource)
      MyBase.New(message)
      Me.ContentSource = contentSource
    End Sub

    Public Sub New(ByVal message As String, ByVal contentSource As ContentSource, ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
      Me.ContentSource = contentSource
    End Sub

#End Region

  End Class

End Namespace