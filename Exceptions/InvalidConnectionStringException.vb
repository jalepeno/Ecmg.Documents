'---------------------------------------------------------------------------------
' <copyright company="ECMG">
'     Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'     Copying or reuse without permission is strictly forbidden.
' </copyright>
'---------------------------------------------------------------------------------


Namespace Exceptions
  ''' <summary>Thrown when invalid connection strings are provided.</summary>
  Public Class InvalidConnectionStringException
    Inherits CtsException

    Public Sub New(ByVal Message As String)
      MyBase.New(Message)
    End Sub

    Public Sub New(ByVal Message As String, ByVal InnerException As Exception)
      MyBase.New(Message, InnerException)
    End Sub

  End Class
End Namespace