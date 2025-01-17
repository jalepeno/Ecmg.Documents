'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Exceptions

  ''' <summary>
  ''' Thrown when a property is requested that does not exist
  ''' </summary>
  ''' <remarks></remarks>
  Public Class PropertyDoesNotExistException
    Inherits CtsException

#Region "Class Variables"

    Private mstrRequestedPropertyName As String = ""

#End Region

#Region "Public Properties"

    Public ReadOnly Property RequestedPropertyName() As String
      Get
        Return mstrRequestedPropertyName
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal lpRequestedPropertyName As String)
      MyBase.New(String.Format("The property '{0}' does not exist.", lpRequestedPropertyName))
    End Sub

    Public Sub New(ByVal message As String, ByVal lpRequestedPropertyName As String)
      Me.New(message, lpRequestedPropertyName, Nothing)
    End Sub

    Public Sub New(ByVal message As String, ByVal lpRequestedPropertyName As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
      mstrRequestedPropertyName = lpRequestedPropertyName
    End Sub

#End Region

  End Class

End Namespace