'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"


Imports System.Xml.Serialization
Imports Documents.Utilities

#End Region

Namespace Core

  ''' <summary>Defines the externally referenced content of a document version.</summary>
  Public Class ReferenceOnlyContent
    Inherits Content

#Region "Public Properties"

    Public Property Uri As String
      Get
        Return mstrContentPath
      End Get
      Set(ByVal value As String)
        ContentPath = value
      End Set
    End Property

    Public Property ContentType As String
      Get
        Return mstrMimeType
      End Get
      Set(ByVal value As String)
        mstrMimeType = value
      End Set
    End Property

#Region "XmlIgnore"

    <XmlIgnore()>
    Public Overrides Property ContentPath() As String
      Get
        If Helper.CallStackContainsMethodName("Serialize") Then
          Return Nothing
        End If
        If mstrContentPath Is Nothing Then
          mstrContentPath = String.Empty
        End If
        Return mstrContentPath
      End Get
      Set(ByVal value As String)
        mstrContentPath = value
      End Set
    End Property

    <XmlIgnore()>
    Public Overrides Property MimeType As String
      Get
        If Helper.CallStackContainsMethodName("Serialize") Then
          Return Nothing
        Else
          Return mstrMimeType
        End If
      End Get
      Set(ByVal value As String)
        mstrMimeType = value
      End Set
    End Property

    <XmlIgnore()>
    Public Shadows Property FileName As String

    <XmlIgnore()>
    Public Shadows Property FileSize As FileSize

    <XmlIgnore()>
    Public Shadows Property Hash As String

    <XmlIgnore()>
    Public Shadows Property Metadata As ECMProperties

    <XmlIgnore()>
    Public Shadows Property RelativePath As String

    <XmlIgnore()>
    Public Overrides Property StorageType() As StorageTypeEnum
      Get
        Return StorageTypeEnum.Reference
      End Get
      Set(ByVal value As StorageTypeEnum)
        ' Ignore, we always want to set to Reference
      End Set
    End Property

    <XmlIgnore()>
    Public Shadows Property Data As ContentData

#End Region

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpVersion As Version)
      MyBase.New(lpVersion)
    End Sub

    Public Sub New(ByVal lpUri As String)
      Uri = lpUri
    End Sub

    Public Sub New(ByVal lpUri As String, ByVal lpContentType As String)
      Uri = lpUri
      ContentType = lpContentType
    End Sub

    Public Sub New(ByVal lpUri As String, ByVal lpVersion As Version)
      MyBase.New(lpVersion)
      Uri = lpUri
    End Sub

    Public Sub New(ByVal lpUri As String, ByVal lpContentType As String, ByVal lpVersion As Version)
      MyBase.New(lpVersion)
      Uri = lpUri
      ContentType = lpContentType
    End Sub

#End Region

#Region "Public Methods"

    Friend Overrides Sub UpdateContentLocation(ByVal lpContentReferenceBehavior As Document.ContentReferenceBehavior)
      '  Do nothing
    End Sub

#End Region

  End Class

End Namespace
