'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "ImportS"

Imports System.ComponentModel
Imports Documents.Providers

#End Region

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the FolderAdded Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FolderAddedEventArgs
    Inherits BackgroundEventArgs

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

    Public Sub New(ByVal lpFolder As IFolder)
      Me.New(lpFolder, Now, Nothing)
    End Sub

    Public Sub New(ByVal lpFolder As IFolder, ByVal lpTime As DateTime)
      Me.New(lpFolder, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpDocument As IFolder, ByVal lpTime As DateTime, ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpTime, lpWorker)
      Folder = lpDocument

    End Sub

    Public Sub New(ByVal lpFolder As IFolder, ByVal lpTime As DateTime, ByVal lpErrorMessage As String, ByVal lpWorker As BackgroundWorker)

      MyBase.New(lpTime, lpErrorMessage, lpWorker)
      Folder = lpFolder

    End Sub

#End Region

  End Class

End Namespace