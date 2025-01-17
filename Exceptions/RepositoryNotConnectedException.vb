'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Namespace Exceptions
  ''' <summary>
  ''' Thrown when the repository connection is not set
  ''' </summary>
  ''' <remarks></remarks>
  Public Class RepositoryNotConnectedException
    Inherits CtsException

#Region "Constructors"

    Public Sub New()
      MyBase.New()
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