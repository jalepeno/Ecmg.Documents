'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  ExportFolderEventArgs.vb
'   Description :  [type_description_here]
'   Created     :  3/6/2015 9:56:20 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  Public Class ExportFolderEventArgs
    Inherits BackgroundWorkerEventArgs

#Region "Class Variables"

    Private mstrPrimaryIdentifier As String = String.Empty
    Private mstrFolderId As String = String.Empty
    Private mstrFolderPath As String = String.Empty
    Private mobjFolder As Folder = Nothing
    Private mblnGetPermissions As Boolean = True

#End Region

#Region "Public Enumerations"

    ''' <summary>
    ''' Denotes whether a folder is referenced by the path or by an id
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum FolderIdentifierType
      ''' <summary>
      ''' Folder is referenced by its id
      ''' </summary>
      ''' <remarks></remarks>
      FolderId = 0
      ''' <summary>
      ''' Folder is referenced by its path
      ''' </summary>
      ''' <remarks></remarks>
      FolderPath = 1
    End Enum

#End Region

#Region "Public Properties"

    Public Property Folder As Folder
      Get
        Try
          Return mobjFolder
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Get
      Set(value As Folder)
        Try
          mobjFolder = value
        Catch ex As Exception
          ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
          ' Re-throw the exception to the caller
          Throw
        End Try
      End Set
    End Property

    Public ReadOnly Property FolderId() As String
      Get
        Return mstrFolderId
      End Get
    End Property

    Public ReadOnly Property FolderPath() As String
      Get
        Return mstrFolderPath
      End Get
    End Property

    ''' <summary>Determines whether or not the export operation should get or exclude the document permissions</summary>
    ''' <remarks>This property can only be used on exports with providers that implement ISecurityClassification.</remarks>
    ''' <seealso cref="Providers.ISecurityClassification">ISecurityClassification</seealso>
    Public Property GetPermissions As Boolean
      Get
        Return mblnGetPermissions
      End Get
      Set(value As Boolean)
        mblnGetPermissions = value
      End Set
    End Property

    Public ReadOnly Property PrimaryIdentifier() As String
      Get
        Return mstrPrimaryIdentifier
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpFolderIdentifier As String, lpFolderIdentifierType As FolderIdentifierType)
      MyBase.New()
      Try
        mstrPrimaryIdentifier = lpFolderIdentifier
        Select Case lpFolderIdentifierType
          Case FolderIdentifierType.FolderId
            mstrFolderId = lpFolderIdentifier
          Case FolderIdentifierType.FolderPath
            mstrFolderPath = lpFolderIdentifier
        End Select
      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod())
        ' Re-throw the exception to the caller
        Throw
      End Try
    End Sub

#End Region

  End Class

End Namespace