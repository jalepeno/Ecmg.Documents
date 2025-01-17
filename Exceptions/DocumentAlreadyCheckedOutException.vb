'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  DocumentAlreadyCheckedOutException.vb
'   Description :  [type_description_here]
'   Created     :  9/10/2013 9:25:20 AM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

#Region "Imports"

#End Region

Namespace Exceptions

  Public Class DocumentAlreadyCheckedOutException
    Inherits DocumentException

#Region "Constructors"

    Public Sub New(ByVal id As String)
      MyBase.New(id, "Document is already checked out.")
    End Sub

    Public Sub New(ByVal id As String, ByVal message As String)
      MyBase.New(id, message)
    End Sub

#End Region

  End Class

End Namespace