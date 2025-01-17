' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------
'  Document    :  IEnumeratedValues.vb
'  Description :  [type_description_here]
'  Created     :  3/15/2012 10:26:04 AM
'  <copyright company="ECMG">
'      Copyright (c) Enterprise Content Management Group, LLC. All rights reserved.
'      Copying or reuse without permission is strictly forbidden.
'  </copyright>
' ---------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------

#Region "Imports"



#End Region

Namespace Core

  Public Interface IEnumeratedValues
    Inherits ICollection(Of IEnumeratedValue)

    Property Item(name As String) As IEnumeratedValue
    Overloads Function Contains(name As String) As Boolean

  End Interface

End Namespace