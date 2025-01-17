' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IPermission.vb.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 3:19:21 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Security

  Public Interface IPermission

    Property PrincipalName As String
    Property Access As IAccessMask
    Property AccessType As AccessType
    Property PrincipalType As PrincipalType
    Property PermissionSource As PermissionSource

  End Interface

End Namespace