' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IPermissionRights.vb
'  Description :  [type_description_here]
'  Created     :  3/16/2012 9:47:13 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"




#End Region

Namespace Security

  Public Interface IPermissionList
    Inherits ICollection(Of PermissionRight)

    Overloads Sub Add(permissionName As String)
    Function ToDelimitedList() As String
    Function FromDelimitedList(lpList As String) As IPermissionList

  End Interface

End Namespace