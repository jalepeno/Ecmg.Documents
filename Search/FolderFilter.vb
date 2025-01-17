'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Search

  ''' <summary>
  ''' Used to limit a search within a folder
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FolderFilter

#Region "Class Variables"

    Private mstrId As String = String.Empty
    Private mstrPath As String = String.Empty
    Private mblnIncludeSubfolders As Boolean

#End Region

#Region "Public Properties"

    Public Property Id() As String
      Get
        Return mstrId
      End Get
      Set(ByVal value As String)
        mstrId = value
      End Set
    End Property

    Public Property Path() As String
      Get
        Return mstrPath
      End Get
      Set(ByVal value As String)
        mstrPath = value
      End Set
    End Property

    Public Property IncludeSubFolders() As Boolean
      Get
        Return mblnIncludeSubfolders
      End Get
      Set(ByVal value As Boolean)
        mblnIncludeSubfolders = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpId As String)
      Me.New(lpId, "", False)
    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpPath As String)
      Me.New(lpId, lpPath, False)
    End Sub

    Public Sub New(ByVal lpId As String, ByVal lpPath As String, ByVal lpIncludeSubFolders As Boolean)
      Id = lpId
      Path = lpPath
      IncludeSubFolders = lpId
    End Sub

#End Region

  End Class

End Namespace