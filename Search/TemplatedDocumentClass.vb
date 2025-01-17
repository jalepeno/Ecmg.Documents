'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"

Imports Documents.Core

#End Region

Namespace Search

  ''' <summary>
  ''' Used in Search Templates for including the way the 
  ''' search user interface is to handle the Document Class
  ''' </summary>
  ''' <remarks></remarks>
  Public Class TemplatedDocumentClass
    Inherits DocumentClass

#Region "Class Variables"

    Private menuView As ElementView

#End Region

#Region "Public Properties"

    Public Property View() As ElementView
      Get
        Return menuView
      End Get
      Set(ByVal value As ElementView)
        menuView = value
      End Set
    End Property
#End Region

#Region "Constructors"

#End Region

  End Class

End Namespace
