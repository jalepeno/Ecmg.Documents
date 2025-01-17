'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Providers

#End Region

Namespace Exceptions
  ''' <summary>
  ''' Exception defined for cases where a 
  ''' property can't be deserialized from a RIF
  ''' </summary>
  ''' <remarks></remarks>
  Public Class PropertyNotInitializedException
    Inherits CtsException

#Region "Class Variables"

    Private mstrProperty As String = String.Empty
    Private mobjDocumentClass As DocumentClass = Nothing
    Private mobjContentSource As ContentSource = Nothing

#End Region

#Region "Public Properties"

    Public Property [Property]() As String
      Get
        Return mstrProperty
      End Get
      Set(ByVal value As String)
        mstrProperty = value
      End Set
    End Property

    Public Property DocumentClass() As DocumentClass
      Get
        Return mobjDocumentClass
      End Get
      Set(ByVal value As DocumentClass)
        mobjDocumentClass = value
      End Set
    End Property

    Public Property ContentSource() As ContentSource
      Get
        Return mobjContentSource
      End Get
      Set(ByVal value As ContentSource)
        mobjContentSource = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Friend Sub New()

    End Sub

    Public Sub New(ByVal [property] As String,
                   ByVal documentClass As DocumentClass,
                   ByVal contentSource As ContentSource)
      MyBase.New(String.Format("The {0} property could not be initialized", [property]))

      With Me
        .[Property] = [property]
        .DocumentClass = documentClass
        .ContentSource = .ContentSource
      End With

    End Sub

    Public Sub New(ByVal message As String,
                   ByVal [property] As String,
                   ByVal documentClass As DocumentClass,
                   ByVal contentSource As ContentSource)
      MyBase.New(message)

      With Me
        .[Property] = [property]
        .DocumentClass = documentClass
        .ContentSource = .ContentSource
      End With

    End Sub

    Public Sub New(ByVal message As String,
                   ByVal [property] As String,
                   ByVal documentClass As DocumentClass,
                   ByVal contentSource As ContentSource, ByVal innerException As System.Exception)
      MyBase.New(message, innerException)

      With Me
        .[Property] = [property]
        .DocumentClass = documentClass
        .ContentSource = .ContentSource
      End With

    End Sub

#End Region

  End Class

End Namespace