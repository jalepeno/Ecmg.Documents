﻿'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Exceptions

  ''' <summary>
  ''' Indicates the specified path is not a valid provider dll.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class InvalidProviderException
    Inherits CtsException

#Region "Class Variables"

    Private m_path As String

#End Region

#Region "Public Properties"

    Public Property Path() As String
      Get
        Return m_path
      End Get
      Set(ByVal value As String)
        m_path = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal path As String)
      MyBase.New(String.Format("The file '{0}' is not a valid provider file.", path))
      Me.Path = path
    End Sub

    Public Sub New(ByVal message As String, ByVal path As String)
      MyBase.New(message)
      Me.Path = path
    End Sub

    Public Sub New(ByVal message As String, ByVal path As String, ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
      Me.Path = path
    End Sub

#End Region

  End Class

End Namespace
