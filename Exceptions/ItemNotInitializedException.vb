' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  ItemNotInitializedException.vb
'  Description :  [type_description_here]
'  Created     :  3/5/2012 1:17:05 PM
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

  Public Class ItemNotInitializedException
    Inherits CtsException

#Region "Class Variables"

    Private mstrItemName As String = String.Empty
    Private mobjContentSource As ContentSource = Nothing

#End Region

#Region "Public Properties"

    Public Property ItemName() As String
      Get
        Return mstrItemName
      End Get
      Set(ByVal value As String)
        mstrItemName = value
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

    Public Sub New()
      MyBase.New()
    End Sub

    Public Sub New(ByVal itemName As String,
                   ByVal contentSource As ContentSource)
      MyBase.New(String.Format("Item '{0}' could not be initialized.", itemName))

      With Me
        .ItemName = itemName
        .ContentSource = contentSource
      End With

    End Sub

    Public Sub New(ByVal message As String,
                   ByVal itemName As String,
                   ByVal contentSource As ContentSource)
      MyBase.New(message)

      With Me
        .ItemName = itemName
        .ContentSource = contentSource
      End With

    End Sub

    Public Sub New(ByVal message As String,
                   ByVal itemName As String,
                   ByVal contentSource As ContentSource,
                   ByVal innerException As System.Exception)
      MyBase.New(message, innerException)

      With Me
        .ItemName = itemName
        .ContentSource = contentSource
      End With

    End Sub

#End Region

  End Class

End Namespace