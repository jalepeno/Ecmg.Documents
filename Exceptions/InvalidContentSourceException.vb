'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Namespace Exceptions
  ''' <summary>Thrown when an invalid content source is provided.</summary>
  Public Class InvalidContentSourceException
    Inherits CtsException

#Region "Constructors"

    Public Sub New()
      MyBase.New("A valid content source is needed to perform this operation.")
    End Sub

    Public Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Public Sub New(ByVal message As String, ByVal innerException As System.Exception)
      MyBase.New(message, innerException)
    End Sub

#End Region

  End Class
End Namespace