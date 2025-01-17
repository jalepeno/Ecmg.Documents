'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  RepositoryAuthenticationException.vb
'   Description :  [type_description_here]
'   Created     :  5/8/2014 9:26:30 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Exceptions

  Public Class RepositoryAuthenticationException
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