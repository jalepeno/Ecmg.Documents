'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Exceptions

  Public Class ApplicationLogNotInitializedException
    Inherits ApplicationException

#Region "Class Variables"

    Private mstrApplicationName As String = String.Empty
    Private mstrLogDirectory As String = String.Empty

#End Region

#Region "Public Properties"

    Public ReadOnly Property ApplicationName() As String
      Get
        Return mstrApplicationName
      End Get
    End Property

    Public ReadOnly Property LogDirectory() As String
      Get
        Return mstrLogDirectory
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal message As String,
                   ByVal applicationName As String,
                   ByVal logDirectory As String)
      MyBase.New(message)
      mstrApplicationName = applicationName
      mstrLogDirectory = logDirectory
    End Sub

    Public Sub New(ByVal message As String,
                   ByVal applicationName As String,
                   ByVal logDirectory As String,
                   ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mstrApplicationName = applicationName
      mstrLogDirectory = logDirectory
    End Sub

#End Region

  End Class

End Namespace