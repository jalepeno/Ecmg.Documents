' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IAccessMask.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 11:07:44 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"




#End Region

Namespace Security

  Public Interface IAccessMask

    Property PermissionList As IPermissionList
    Property Name As String
    Property Value As Nullable(Of Integer)
    ReadOnly Property Type As PermissionType

  End Interface

End Namespace
