'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

#Region "Imports"




#End Region

''' <summary>
''' Contains the defined exceptions of the Cts framework
''' </summary>
Namespace Exceptions

  ''' <summary>
  ''' Exception thrown when attempting to access a document reference 
  ''' from a transformation when the document property has not been set
  ''' </summary>
  ''' <remarks></remarks>
  Public Class DocumentReferenceNotSetException
    Inherits CtsException

#Region "Private Constants"

    Private Const DEFAULT_MESSAGE As String = "The document reference has not been set"

#End Region

#Region "Constructors"

    Public Sub New()
      MyBase.New(DEFAULT_MESSAGE)
    End Sub

    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
    End Sub

#End Region

  End Class

End Namespace