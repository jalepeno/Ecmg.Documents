'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  DocumentNotCheckedOutException.vb
'   Description :  [type_description_here]
'   Created     :  9/10/2013 9:28:27 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Exceptions

  Public Class DocumentNotCheckedOutException
    Inherits DocumentException

#Region "Constructors"

    Public Sub New(ByVal id As String)
      MyBase.New(id, "Document is not checked out.")
    End Sub

    Public Sub New(ByVal id As String, ByVal message As String)
      MyBase.New(id, message)
    End Sub

#End Region

  End Class

End Namespace