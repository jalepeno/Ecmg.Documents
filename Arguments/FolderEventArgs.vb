'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "ImportS"

Imports Documents.Providers

#End Region

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for a Folder based Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FolderEventArgs
    Inherits EventArgs

#Region "Class Variables"

    Private mobjFolder As IFolder

#End Region

#Region "Public Properties"

    Public Property Folder() As IFolder
      Get
        Return mobjFolder
      End Get
      Set(ByVal value As IFolder)
        mobjFolder = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpFolder As IFolder)
      MyBase.New()
      mobjFolder = lpFolder
    End Sub

#End Region

  End Class

End Namespace