'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Base class for events DocumentFilng Events
  ''' </summary>
  ''' <remarks></remarks>
  Public MustInherit Class DocumentFilingEventArgs
    Inherits DocumentEventArgs

#Region "Class Variables"

    Private mstrId As String
    Private mstrFolderPath As String

#End Region

#Region "Public Properties"

    Public ReadOnly Property Id() As String
      Get
        Return mstrId
      End Get
    End Property

    Public ReadOnly Property FolderPath() As String
      Get
        Return mstrFolderPath
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpId As String, ByVal lpFolderPath As String)
      Me.New(lpId, lpFolderPath, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpFolderPath As String)
      Me.New(lpDocument, lpFolderPath, Now)
    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpFolderPath As String, ByVal lpTime As DateTime)
      MyBase.New(String.Empty, lpTime)
      mstrId = lpId
      mstrFolderPath = lpFolderPath
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpFolderPath As String, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, lpTime)
      mstrId = lpDocument.ID
      mstrFolderPath = lpFolderPath
    End Sub

#End Region

  End Class
End Namespace