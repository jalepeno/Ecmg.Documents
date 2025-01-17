'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments

  Public Class UpdateSystemPropertiesArgs

#Region "Class Variables"

    Private mstrTargetId As String = String.Empty
    Private mobjTarget As IMetaHolder
    Private mstrSystemVersion As String = String.Empty

#End Region

#Region "Public Properties"

    Public Property TargetId() As String
      Get
        Return mstrTargetId
      End Get
      Set(ByVal value As String)
        mstrTargetId = value
      End Set
    End Property

    Public Property Target() As IMetaHolder
      Get
        Return mobjTarget
      End Get
      Set(ByVal value As IMetaHolder)
        mobjTarget = value
      End Set
    End Property

    Public Property SystemVersion() As String
      Get
        Return mstrSystemVersion
      End Get
      Set(ByVal value As String)
        mstrSystemVersion = value
      End Set
    End Property

#End Region

#Region "Constructors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal lpObjectId As String, ByVal lpTarget As IMetaHolder)
      TargetId = lpObjectId
      Target = lpTarget
    End Sub

    Public Sub New(ByVal lpObjectId As String, ByVal lpTarget As IMetaHolder, ByVal lpSystemVersion As String)
      TargetId = lpObjectId
      Target = lpTarget
      SystemVersion = lpSystemVersion
    End Sub

#End Region

  End Class

End Namespace
