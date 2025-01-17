' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IAccessLevel.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 1:11:03 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"




#End Region

Namespace Security

  Public Interface IAccessLevel
    Inherits IAccessMask

    Property Level As PermissionLevel
    Property Rights As IAccessRights

  End Interface

End Namespace