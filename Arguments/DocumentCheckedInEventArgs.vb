'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  ''' <summary>
  ''' Contains all the parameters for the DocumentCheckedIn Event
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentCheckedInEventArgs
    Inherits DocumentCheckedOutEventArgs

#Region "Class Variables"

    Private mblnAsMajorVersion As Boolean
    Private mstrContentPaths As New List(Of String)

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' Indicates whether or not the document was checked in as a major version
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property AsMajorVersion() As Boolean
      Get
        Return mblnAsMajorVersion
      End Get
      Set(ByVal value As Boolean)
        mblnAsMajorVersion = value
      End Set
    End Property

    ''' <summary>
    ''' The current version id of the document after checkin
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Property VersionId() As String
      Get
        Return MyBase.VersionId
      End Get
      Set(ByVal value As String)
        MyBase.VersionId = value
      End Set
    End Property

    Public ReadOnly Property ContentPath() As String
      Get
        Return ContentPaths(0)
      End Get
    End Property

    Public ReadOnly Property ContentPaths() As List(Of String)
      Get
        Return mstrContentPaths
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document, ByVal lpContentPath As String, ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String, ByVal lpAsMajorVersion As Boolean)
      Me.New(lpDocument, lpContentPath, lpCheckedOutUserName, lpVersionId, lpAsMajorVersion, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpContentPath As String, ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String, ByVal lpAsMajorVersion As Boolean, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, lpContentPath, New String() {}, lpCheckedOutUserName, lpVersionId, lpTime)
      AsMajorVersion = lpAsMajorVersion
      EventDescription = "DocumentCheckedIn"
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpContentPaths() As String, ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String, ByVal lpAsMajorVersion As Boolean)
      Me.New(lpDocument, lpContentPaths, lpCheckedOutUserName, lpVersionId, lpAsMajorVersion, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpContentPaths() As String, ByVal lpCheckedOutUserName As String, ByVal lpVersionId As String, ByVal lpAsMajorVersion As Boolean, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, "", lpContentPaths, lpCheckedOutUserName, lpVersionId, lpTime)
      AsMajorVersion = lpAsMajorVersion
      For lintContentPathCounter As Integer = 0 To lpContentPaths.Length - 1
        mstrContentPaths.Add(lpContentPaths(lintContentPathCounter))
      Next
      EventDescription = "DocumentCheckedIn"
    End Sub

#End Region

  End Class
End Namespace