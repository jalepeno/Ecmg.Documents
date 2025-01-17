'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentImported Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentAddedEventArgs
    Inherits DocumentEventArgs

#Region "Class Variables"

    Private mstrNewId As String

#End Region

#Region "Public Properties"

    Public ReadOnly Property NewId() As String
      Get
        Return mstrNewId
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document, ByVal lpNewId As String)
      Me.New(lpDocument, lpNewId, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpNewId As String, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, "DocumentAdded", lpTime)
      mstrNewId = lpNewId
    End Sub

#End Region

  End Class
End Namespace