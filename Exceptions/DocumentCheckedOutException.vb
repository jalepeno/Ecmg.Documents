'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  DocumentCheckedOutException.vb
'   Description :  [type_description_here]
'   Created     :  4/29/2014 1:48:45 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Exceptions

  Public Class DocumentCheckedOutException
    Inherits DocumentException

#Region "Constructors"

    Public Sub New(ByVal id As String)
      MyBase.New(id, "Document is checked out.")
    End Sub

    Public Sub New(ByVal id As String, ByVal message As String)
      MyBase.New(id, message)
    End Sub

#End Region

  End Class

End Namespace