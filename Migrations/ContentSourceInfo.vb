'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports System.Xml.Serialization

Namespace Migrations
  ''' <summary>
  ''' Describes a ContentSource object in the context of a
  ''' CondensedMigrationConfiguration.
  ''' </summary>
  Public NotInheritable Class ContentSourceInfo

#Region "Class Variables"

    Private mstrName As String
    Private mstrConnectionString As String

#End Region

#Region "Public Properties"

    <XmlAttribute("name")>
    Public Property Name() As String
      Get
        Return mstrName
      End Get
      Set(ByVal value As String)
        mstrName = value
      End Set
    End Property

    Public Property ConnectionString() As String
      Get
        Return mstrConnectionString
      End Get
      Set(ByVal value As String)
        mstrConnectionString = value
      End Set
    End Property

#End Region

  End Class
End Namespace