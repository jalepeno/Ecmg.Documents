'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentCheckedOut Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentCheckedOutEventArgs
    Inherits DocumentCopiedOutEventArgs

#Region "Class Variables"

    Private mstrCheckedOutUserName As String = ""
    Private mstrVersionId As String = ""

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The UserName of the user that checked out the document
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CheckedOutUserName() As String
      Get
        Return mstrCheckedOutUserName
      End Get
      Set(ByVal value As String)
        mstrCheckedOutUserName = value
      End Set
    End Property

    ''' <summary>
    ''' The current version id of the document at the time it was checked out
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property VersionId() As String
      Get
        Return mstrVersionId
      End Get
      Set(ByVal value As String)
        mstrVersionId = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document, ByVal lpDestinationPath As String, ByVal lpOutputFileNames As String(), ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String)
      Me.New(lpDocument, lpDestinationPath, lpOutputFileNames, lpCheckedOutUserName, lpVersionId, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpDestinationPath As String, ByVal lpOutputFileNames As String(), ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, lpDestinationPath, lpOutputFileNames, lpTime)
      CheckedOutUserName = lpCheckedOutUserName
      VersionId = lpVersionId
      EventDescription = "DocumentCheckedOut"
    End Sub

#End Region

  End Class
End Namespace