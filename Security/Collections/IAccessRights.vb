' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IAccessRights.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 1:12:44 PM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"




#End Region

Namespace Security

  Public Interface IAccessRights
    Inherits ICollection(Of IAccessMask)

    Property Item(name As String) As IAccessMask
    ReadOnly Property Item(maskValue As Integer) As IAccessMask
    Overloads Function Contains(name As String) As Boolean

  End Interface

End Namespace