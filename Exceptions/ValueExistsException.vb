'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------

Namespace Exceptions

  Public Class ValueExistsException
    Inherits CtsException

#Region "Class Variables"

    Private mobjValue As String

#End Region

#Region "Public Properties"

    ''' <summary>
    ''' The existing value
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Value() As String
      Get
        Return mobjValue
      End Get
    End Property

#End Region

#Region "Constructors"

    Public Sub New(ByVal value As String,
                   ByVal message As String)
      MyBase.New(message)
      mobjValue = value
    End Sub

    ''' <summary>
    ''' Creates a new ValueAlreadyExistsException with the specified value.
    ''' </summary>
    ''' <param name="value">The existing value</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal value As String,
                   ByVal message As String,
                   ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
      mobjValue = value
    End Sub

#End Region

  End Class

End Namespace