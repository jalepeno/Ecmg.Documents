' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  FolderPropertyArgs.vb
'  Description :  [type_description_here]
'  Created     :  2/29/2012 1:41:43 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core
Imports Documents.Utilities

#End Region

Namespace Arguments

  ''' <summary>
  ''' Contains all the necessary parameters for the UpdateFolder method
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FolderPropertyArgs
    Inherits ObjectPropertyArgs

#Region "Class Variables"

    Private mstrFolderPath As String = String.Empty

#End Region

#Region "Public Properties"

    Public Property FolderId As String
      Get
        Return ObjectID
      End Get
      Set(value As String)
        ObjectID = value
      End Set
    End Property

    Public Property FolderPath As String
      Get
        Return mstrFolderPath
      End Get
      Set(value As String)
        mstrFolderPath = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    ''' <summary>
    ''' Creates a new instance of ObjectPropertyArgs
    ''' </summary>
    ''' <param name="lpFolderId">The folder identifier</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpFolderId As String)
      Me.New(lpFolderId, String.Empty, New ECMProperties)
    End Sub

    ''' <summary>
    ''' Creates a new instance of ObjectPropertyArgs
    ''' </summary>
    ''' <param name="lpFolderId">The folder identifier</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpFolderId As String, lpFolderPath As String)
      Me.New(lpFolderId, lpFolderPath, New ECMProperties)
    End Sub

    ''' <summary>
    ''' Creates a new instance of DocumentUpdateArgs
    ''' </summary>
    ''' <param name="lpFolderId">The folder identifier</param>
    ''' <param name="lpProperties">A Core.Properties object containing the properties to update</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal lpFolderId As String, lpFolderPath As String, ByVal lpProperties As IProperties)

      Try

        FolderId = lpFolderId
        FolderPath = lpFolderPath
        Properties = lpProperties

      Catch ex As Exception
        ApplicationLogging.LogException(ex, Reflection.MethodBase.GetCurrentMethod)
      End Try

    End Sub

#End Region

  End Class

End Namespace