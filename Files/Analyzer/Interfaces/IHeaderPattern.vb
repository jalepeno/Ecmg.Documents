'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------
'   Document    :  IHeaderPattern.vb
'   Description :  [type_description_here]
'   Created     :  1/26/2015 1:56:19 PM
'   <copyright company="ECMG">
'       Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'       Copying or reuse without permission is strictly forbidden.
'   </copyright>
'  ---------------------------------------------------------------------------------
'  ---------------------------------------------------------------------------------

Namespace Files
  Public Interface IHeaderPattern

    Property Bytes As String
    Property Position As Integer
    ReadOnly Property Length As Integer
    ReadOnly Property Pattern As Byte()
    ReadOnly Property Points As Integer
    ReadOnly Property XM As Boolean

  End Interface

End Namespace