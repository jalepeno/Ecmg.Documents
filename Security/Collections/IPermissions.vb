' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IPermissions.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 3:52:32 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"





#End Region

Namespace Security

  Public Interface IPermissions
    Inherits ICollection(Of IPermission)
    Sub AddRange(permissions As IPermissions)
    Property Item(name As String) As IPermission
    Overloads Function Contains(name As String) As Boolean
  End Interface

End Namespace