' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IEnumeratedValue.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 10:22:37 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

Namespace Core

  Public Interface IEnumeratedValue

    Property ParentName As String
    Property Name As String
    Property Value As Nullable(Of Integer)

  End Interface

End Namespace