'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  VersioningNotEnabledException.vb
'   Description :  [type_description_here]
'   Created     :  5/20/2015 9:23:02 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Exceptions

  Public Class VersioningNotEnabledException
    Inherits CtsException

#Region "Constructors"

    Public Sub New()
      MyBase.New("Versioning is not enabled for class or repository.")
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