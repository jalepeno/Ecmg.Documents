'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Imports Documents.Core

Namespace Arguments
  Public Class DocumentUpdatedEventArgs
    Inherits DocumentEventArgs

#Region "Class Variables"

    Private mobjUpdatedProperties As IProperties

#End Region

#Region "Public Properties"

    Public ReadOnly Property UpdatedProperties() As IProperties
      Get
        Return mobjUpdatedProperties
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpDocument As Document, ByVal lpDocumentPropertyArgs As DocumentPropertyArgs)
      Me.New(lpDocument, lpDocumentPropertyArgs.Properties)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpUpdatedProperties As IProperties)
      Me.New(lpDocument, lpUpdatedProperties, Now)
    End Sub

    Public Sub New(ByVal lpDocument As Document, ByVal lpUpdatedProperties As IProperties, ByVal lpTime As DateTime)
      MyBase.New(lpDocument, "DocumentUpdated", lpTime)
      mobjUpdatedProperties = lpUpdatedProperties
    End Sub

#End Region

  End Class

End Namespace