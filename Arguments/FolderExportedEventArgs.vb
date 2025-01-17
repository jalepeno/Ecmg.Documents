'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.ComponentModel

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the FolderExported Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FolderExportedEventArgs
    Inherits BackgroundEventArgs

#Region "Class Variables"

    Private mobjFolder As Providers.IFolder
    Private mstrExportPath As String = ""

#End Region

#Region "Public Properties"

    Public ReadOnly Property Folder() As Providers.IFolder
      Get
        Return mobjFolder
      End Get
    End Property

    Public ReadOnly Property ExportPath() As String
      Get
        Return mstrExportPath
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpFolder As Providers.IFolder,
                   ByVal lpExportPath As String,
                   ByVal lpTime As DateTime)
      Me.New(lpFolder, lpExportPath, lpTime, Nothing)
    End Sub

    Public Sub New(ByVal lpFolder As Providers.IFolder,
                   ByVal lpExportPath As String,
                   ByVal lpTime As DateTime,
                   ByVal lpWorker As BackgroundWorker)
      MyBase.New(lpTime, lpWorker)
      mobjFolder = lpFolder
      mstrExportPath = lpExportPath
    End Sub

#End Region

  End Class
End Namespace